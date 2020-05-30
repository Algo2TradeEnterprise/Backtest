Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class LossMakeupFavourableFractalBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public MaxProfitPerTrade As Decimal
        Public MinimumTargetATRMultipler As Decimal
        Public MinimumStoplossATRMultipler As Decimal
        Public MaximumStoplossATRMultipler As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _fractalHighPayload As Dictionary(Of Date, Decimal)
    Private _fractalLowPayload As Dictionary(Of Date, Decimal)
    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _previousDayHighestATR As Decimal
    Private _considerPreviousDayFractalForBuy As Boolean = False
    Private _considerPreviousDayFractalForSell As Boolean = False
    Private _lastPotentialPL As Decimal = 0
    Private _lastCancelOrder As Trade = Nothing

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)

        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            Dim firstCandle As Payload = Nothing
            For Each runningPayload In _signalPayload
                If runningPayload.Key.Date = _tradingDate.Date Then
                    If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                        firstCandle = runningPayload.Value
                        Exit For
                    End If
                End If
            Next
            If firstCandle IsNot Nothing AndAlso firstCandle.PreviousCandlePayload IsNot Nothing Then
                If firstCandle.Open > _fractalHighPayload(firstCandle.PayloadDate) AndAlso firstCandle.Open > _fractalLowPayload(firstCandle.PayloadDate) Then
                    _considerPreviousDayFractalForSell = True
                End If
                If firstCandle.Open < _fractalHighPayload(firstCandle.PayloadDate) AndAlso firstCandle.Open < _fractalLowPayload(firstCandle.PayloadDate) Then
                    _considerPreviousDayFractalForBuy = True
                End If

                _previousDayHighestATR = _atrPayload.Max(Function(x)
                                                             If x.Key.Date = firstCandle.PreviousCandlePayload.PayloadDate.Date Then
                                                                 Return x.Value
                                                             Else
                                                                 Return Decimal.MinValue
                                                             End If
                                                         End Function)
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signal As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 AndAlso (lastExecutedOrder Is Nothing OrElse Not IsBothFractalSame(_fractalHighPayload, currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, lastExecutedOrder.SignalCandle.PayloadDate)) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim triggerPrice As Decimal = signal.Item2 + buffer
                    Dim stoplossPrice As Decimal = signal.Item2 - ConvertFloorCeling(_previousDayHighestATR, _parentStrategy.TickSize, RoundOfType.Celing) - buffer
                    'Dim stoplossPrice As Decimal = _fractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) - buffer
                    'Dim slRemark As String = "Fractal"
                    'Dim atr As Decimal = ConvertFloorCeling(_atrPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                    'Dim minATRSL As Decimal = ConvertFloorCeling(atr * _userInputs.MinimumStoplossATRMultipler, _parentStrategy.TickSize, RoundOfType.Celing)
                    'Dim maxATRSL As Decimal = ConvertFloorCeling(atr * _userInputs.MaximumStoplossATRMultipler, _parentStrategy.TickSize, RoundOfType.Celing)
                    'Dim slPoint As Decimal = (triggerPrice - stoplossPrice) - 2 * buffer
                    'If slPoint < minATRSL Then
                    '    stoplossPrice = triggerPrice - minATRSL - buffer
                    '    slRemark = "Min ATR SL"
                    'ElseIf slPoint > maxATRSL Then
                    '    stoplossPrice = triggerPrice - maxATRSL - buffer
                    '    slRemark = "Max ATR SL"
                    'End If
                    If stoplossPrice < triggerPrice Then
                        Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, triggerPrice, stoplossPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, _parentStrategy.StockType)
                        Dim targetPrice As Decimal = Decimal.MaxValue
                        If _userInputs.MaxProfitPerTrade = Decimal.MaxValue Then
                            targetPrice = triggerPrice + 1000000
                        Else
                            targetPrice = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, triggerPrice, quantity, _userInputs.MaxProfitPerTrade, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                        End If

                        parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = triggerPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoplossPrice,
                                    .Target = targetPrice,
                                    .Buffer = buffer,
                                    .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = "Normal",
                                    .Supporting3 = Math.Abs(triggerPrice - stoplossPrice)
                                }

                        Dim pl As Decimal = GetStockPotentialPL(currentMinuteCandlePayload)
                        If pl < 0 Then
                            Dim targetPoint As Decimal = ConvertFloorCeling((triggerPrice - stoplossPrice) / 2, _parentStrategy.TickSize, RoundOfType.Celing)
                            'Dim atrTarget As Decimal = ConvertFloorCeling(atr * _userInputs.MinimumTargetATRMultipler, _parentStrategy.TickSize, RoundOfType.Celing)
                            Dim atrTarget As Decimal = ConvertFloorCeling(_previousDayHighestATR * _userInputs.MinimumTargetATRMultipler, _parentStrategy.TickSize, RoundOfType.Celing) + 2 * buffer
                            targetPoint = Math.Max(atrTarget, targetPoint)
                            targetPrice = triggerPrice + targetPoint
                            quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, triggerPrice, targetPrice, Math.Abs(pl), Trade.TypeOfStock.Cash)

                            parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = triggerPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoplossPrice,
                                    .Target = targetPrice,
                                    .Buffer = buffer,
                                    .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = "SL Makeup Trade",
                                    .Supporting3 = Math.Abs(triggerPrice - stoplossPrice),
                                    .Supporting4 = pl
                                }
                        End If
                    End If
                End If
            End If
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signal As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 AndAlso (lastExecutedOrder Is Nothing OrElse Not IsBothFractalSame(_fractalLowPayload, currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, lastExecutedOrder.SignalCandle.PayloadDate)) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim triggerPrice As Decimal = signal.Item2 - buffer
                    Dim stoplossPrice As Decimal = signal.Item2 + ConvertFloorCeling(_previousDayHighestATR, _parentStrategy.TickSize, RoundOfType.Celing) + buffer
                    'Dim stoplossPrice As Decimal = _fractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) + buffer
                    'Dim slRemark As String = "Fractal"
                    'Dim atr As Decimal = ConvertFloorCeling(_atrPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                    'Dim minATRSL As Decimal = ConvertFloorCeling(atr * _userInputs.MinimumStoplossATRMultipler, _parentStrategy.TickSize, RoundOfType.Celing)
                    'Dim maxATRSL As Decimal = ConvertFloorCeling(atr * _userInputs.MaximumStoplossATRMultipler, _parentStrategy.TickSize, RoundOfType.Celing)
                    'Dim slPoint As Decimal = (stoplossPrice - triggerPrice) - 2 * buffer
                    'If slPoint < minATRSL Then
                    '    stoplossPrice = triggerPrice + minATRSL + buffer
                    '    slRemark = "Min ATR SL"
                    'ElseIf slPoint > maxATRSL Then
                    '    stoplossPrice = triggerPrice + maxATRSL + buffer
                    '    slRemark = "Max ATR SL"
                    'End If
                    If stoplossPrice > triggerPrice Then
                        Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoplossPrice, triggerPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, _parentStrategy.StockType)
                        Dim targetPrice As Decimal = Decimal.MinValue
                        If _userInputs.MaxProfitPerTrade = Decimal.MaxValue Then
                            targetPrice = triggerPrice - 1000000
                        Else
                            targetPrice = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, triggerPrice, quantity, _userInputs.MaxProfitPerTrade, Trade.TradeExecutionDirection.Sell, _parentStrategy.StockType)
                        End If

                        parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = triggerPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoplossPrice,
                                    .Target = targetPrice,
                                    .Buffer = buffer,
                                    .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = "Normal",
                                    .Supporting3 = Math.Abs(triggerPrice - stoplossPrice)
                                }

                        Dim pl As Decimal = GetStockPotentialPL(currentMinuteCandlePayload)
                        If pl < 0 Then
                            Dim targetPoint As Decimal = ConvertFloorCeling((stoplossPrice - triggerPrice) / 2, _parentStrategy.TickSize, RoundOfType.Celing)
                            'Dim atrTarget As Decimal = ConvertFloorCeling(atr * _userInputs.MinimumTargetATRMultipler, _parentStrategy.TickSize, RoundOfType.Celing)
                            Dim atrTarget As Decimal = ConvertFloorCeling(_previousDayHighestATR * _userInputs.MinimumTargetATRMultipler, _parentStrategy.TickSize, RoundOfType.Celing) + 2 * buffer
                            targetPoint = Math.Max(atrTarget, targetPoint)
                            targetPrice = triggerPrice - targetPoint
                            quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, targetPrice, triggerPrice, Math.Abs(pl), Trade.TypeOfStock.Cash)

                            parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = triggerPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoplossPrice,
                                    .Target = targetPrice,
                                    .Buffer = buffer,
                                    .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = "SL Makeup Trade",
                                    .Supporting3 = Math.Abs(triggerPrice - stoplossPrice),
                                    .Supporting4 = pl
                                }
                        End If
                    End If
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter1)
        End If
        If parameter2 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter2)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            If _lastCancelOrder IsNot Nothing AndAlso _lastCancelOrder.EntryTime = currentTrade.EntryTime Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            Else
                Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                Dim pl As Decimal = GetStockPotentialPL(currentMinuteCandlePayload)
                If _lastPotentialPL <> pl Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    _lastCancelOrder = currentTrade
                Else
                    Dim signal As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, currentTrade.EntryDirection)
                    If signal IsNot Nothing Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                        Dim entryPrice As Decimal = Decimal.MinValue
                        If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                            entryPrice = signal.Item2 + buffer
                        ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                            entryPrice = signal.Item2 - buffer
                        End If
                        If entryPrice <> Decimal.MinValue AndAlso entryPrice <> currentTrade.EntryPrice Then
                            ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                        End If
                    End If
                End If
                _lastPotentialPL = pl
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim fractal As Decimal = _fractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                If fractal - currentTrade.StoplossBuffer > currentTrade.PotentialStopLoss Then
                    If Not IsBothFractalSame(_fractalLowPayload, currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, currentTrade.SignalCandle.PayloadDate) Then
                        triggerPrice = fractal - currentTrade.StoplossBuffer
                    End If
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim fractal As Decimal = _fractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                If fractal + currentTrade.StoplossBuffer < currentTrade.PotentialStopLoss Then
                    If Not IsBothFractalSame(_fractalHighPayload, currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, currentTrade.SignalCandle.PayloadDate) Then
                        triggerPrice = fractal + currentTrade.StoplossBuffer
                    End If
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Moved to {0} at {1}", triggerPrice, currentTick.PayloadDate.ToString("HH:mm:ss")))
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Decimal)
        Dim ret As Tuple(Of Boolean, Decimal) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If direction = Trade.TradeExecutionDirection.Buy Then
                Dim lastfractal As Decimal = GetLastDifferentFractal(_fractalHighPayload, candle.PayloadDate, _considerPreviousDayFractalForBuy)
                If lastfractal <> Decimal.MinValue AndAlso _fractalHighPayload(candle.PayloadDate) < lastfractal Then
                    ret = New Tuple(Of Boolean, Decimal)(True, _fractalHighPayload(candle.PayloadDate))
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                Dim lastfractal As Decimal = GetLastDifferentFractal(_fractalLowPayload, candle.PayloadDate, _considerPreviousDayFractalForSell)
                If lastfractal <> Decimal.MinValue AndAlso _fractalLowPayload(candle.PayloadDate) > lastfractal Then
                    ret = New Tuple(Of Boolean, Decimal)(True, _fractalLowPayload(candle.PayloadDate))
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetLastDifferentFractal(ByVal fractalPayload As Dictionary(Of Date, Decimal), ByVal currentTime As Date, ByVal considerPreviousDay As Boolean) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If fractalPayload IsNot Nothing AndAlso fractalPayload.Count > 0 Then

            For Each runningFractal In fractalPayload.OrderByDescending(Function(x)
                                                                            Return x.Key
                                                                        End Function)
                If runningFractal.Key < currentTime AndAlso (considerPreviousDay OrElse runningFractal.Key.Date = _tradingDate.Date) Then
                    If runningFractal.Value <> fractalPayload(currentTime) Then
                        ret = runningFractal.Value
                        Exit For
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Private Function IsBothFractalSame(ByVal fractalPayload As Dictionary(Of Date, Decimal), ByVal currentSignalTime As Date, ByVal lastSignalTime As Date) As Boolean
        Dim ret As Boolean = False
        If fractalPayload IsNot Nothing AndAlso fractalPayload.Count > 0 Then
            Dim currentSignalFractal As Decimal = fractalPayload(currentSignalTime)
            Dim lastSignalFractal As Decimal = fractalPayload(lastSignalTime)
            If currentSignalFractal = lastSignalFractal Then
                For Each runningFractal In fractalPayload.OrderByDescending(Function(x)
                                                                                Return x.Key
                                                                            End Function)
                    If runningFractal.Key < currentSignalTime AndAlso runningFractal.Key >= lastSignalTime Then
                        If runningFractal.Value <> currentSignalFractal Then
                            ret = False
                            Exit For
                        ElseIf runningFractal.Key = lastSignalTime Then
                            ret = True
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        Return ret
    End Function

    Private Function GetStockPotentialPL(ByVal currentMinutePayload As Payload) As Decimal
        Dim ret As Decimal = 0
        Dim completeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
        If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then
            ret += completeTrades.Sum(Function(x)
                                          Return x.PLAfterBrokerage
                                      End Function)
        End If
        Dim runningTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Inprogress)
        If runningTrades IsNot Nothing AndAlso runningTrades.Count > 0 Then
            For Each runningTrade In runningTrades
                If runningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    ret += _parentStrategy.CalculatePL(runningTrade.TradingSymbol, runningTrade.EntryPrice, runningTrade.PotentialStopLoss, runningTrade.Quantity, runningTrade.LotSize, _parentStrategy.StockType)
                ElseIf runningTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    ret += _parentStrategy.CalculatePL(runningTrade.TradingSymbol, runningTrade.PotentialStopLoss, runningTrade.EntryPrice, runningTrade.Quantity, runningTrade.LotSize, _parentStrategy.StockType)
                End If
            Next
        End If
        Return ret
    End Function
End Class
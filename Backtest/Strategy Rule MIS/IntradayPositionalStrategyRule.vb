
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class IntradayPositionalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxStoplossPerStock As Decimal
        Public TargetMultiplier As Decimal
        Public BreakevenMovement As Boolean
        Public LowSLTargetMultiplier As Decimal
        Public LowSLMaxTarget As Decimal
    End Class
#End Region

    Private _fractalHighPayload As Dictionary(Of Date, Decimal)
    Private _fractalLowPayload As Dictionary(Of Date, Decimal)

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _direction As Integer
    Private ReadOnly _dayATR As Decimal

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal direction As Integer,
                   ByVal dayATR As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _direction = direction
        _dayATR = dayATR
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload, currentTick)
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If lastExecutedTrade Is Nothing Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                ElseIf lastExecutedTrade.SignalCandle.PayloadDate <> currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                End If
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandleSatisfied.Item2 + buffer
                    Dim slPrice As Decimal = signalCandleSatisfied.Item3 - buffer
                    Dim slPoint As Decimal = entryPrice - slPrice
                    Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                    If slPoint < _dayATR / 4 Then
                        Dim quantity As Decimal = 1
                        Dim slRemark As String = "Normal SL"
                        If slPoint <= Math.Round(_dayATR / 8, 1) Then
                            slPrice = entryPrice - ConvertFloorCeling(_dayATR / 8, _parentStrategy.TickSize, RoundOfType.Floor)
                            slRemark = "1/8 Of Day ATR"

                            slPoint = entryPrice - slPrice
                            targetPoint = ConvertFloorCeling(slPoint * _userInputs.LowSLTargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                            quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice + targetPoint, _userInputs.LowSLMaxTarget, Trade.TypeOfStock.Cash)
                        Else
                            slPoint = entryPrice - slPrice
                            quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, slPrice, Math.Abs(_userInputs.MaxStoplossPerStock) * -1, Trade.TypeOfStock.Cash)
                        End If

                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = slPrice,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = slRemark,
                                    .Supporting3 = slPrice
                                }
                    End If
                ElseIf signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandleSatisfied.Item2 - buffer
                    Dim slPrice As Decimal = signalCandleSatisfied.Item3 + buffer
                    Dim slPoint As Decimal = slPrice - entryPrice
                    Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                    If slPoint < _dayATR / 4 Then
                        Dim quantity As Decimal = 1
                        Dim slRemark As String = "Normal SL"
                        If slPoint <= Math.Round(_dayATR / 8, 1) Then
                            slPrice = entryPrice + ConvertFloorCeling(_dayATR / 8, _parentStrategy.TickSize, RoundOfType.Floor)
                            slRemark = "1/8 of Day ATR"

                            slPoint = slPrice - entryPrice
                            targetPoint = ConvertFloorCeling(slPoint * _userInputs.LowSLTargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                            quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice - targetPoint, entryPrice, _userInputs.LowSLMaxTarget, Trade.TypeOfStock.Cash)
                        Else
                            slPoint = slPrice - entryPrice
                            quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, slPrice, entryPrice, Math.Abs(_userInputs.MaxStoplossPerStock) * -1, Trade.TypeOfStock.Cash)
                        End If

                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = slPrice,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = slRemark,
                                    .Supporting3 = slPrice
                                }
                    End If
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        'Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        'If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
        '    Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
        '    If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
        '        If currentTrade.SignalCandle.PayloadDate <> currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
        '            ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
        '        End If
        '    End If
        'End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If _userInputs.BreakevenMovement Then
                Dim triggerPrice As Decimal = Decimal.MinValue
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    Dim excpectedTarget As Decimal = currentTrade.EntryPrice + (currentTrade.EntryPrice - currentTrade.PotentialStopLoss) * 1.1
                    If currentTick.Open >= excpectedTarget Then
                        triggerPrice = currentTrade.EntryPrice + _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, LotSize, _parentStrategy.StockType)
                    End If
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    Dim excpectedTarget As Decimal = currentTrade.EntryPrice - (currentTrade.PotentialStopLoss - currentTrade.EntryPrice) * 1.1
                    If currentTick.Open <= excpectedTarget Then
                        triggerPrice = currentTrade.EntryPrice - _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, LotSize, _parentStrategy.StockType)
                    End If
                End If
                If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Move to breakeven: {0}. Time:{1}", triggerPrice, currentTick.PayloadDate))
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Private Function GetSignalCandle(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection) = Nothing
        If currentCandle.PreviousCandlePayload IsNot Nothing AndAlso Not currentCandle.PreviousCandlePayload.DeadCandle Then
            If _direction > 0 Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(_fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate), RoundOfType.Floor)
                If currentCandle.High >= _fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) + buffer AndAlso IsFractalHighSatisfied(currentCandle) Then
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate), _fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate), Trade.TradeExecutionDirection.Buy)
                End If
            ElseIf _direction < 0 Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(_fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate), RoundOfType.Floor)
                If currentCandle.Low <= _fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) AndAlso IsFractalLowSatisfied(currentCandle) Then
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate), _fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate), Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsFractalHighSatisfied(ByVal currentCandle As Payload) As Boolean
        Dim ret As Boolean = False
        For Each runningPayload In _signalPayload.OrderByDescending(Function(x)
                                                                        Return x.Key
                                                                    End Function)
            If runningPayload.Key < currentCandle.PayloadDate AndAlso runningPayload.Key.Date = _tradingDate.Date Then
                If _fractalHighPayload(runningPayload.Value.PayloadDate) < _fractalHighPayload(runningPayload.Value.PreviousCandlePayload.PayloadDate) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(_fractalHighPayload(runningPayload.Value.PayloadDate), RoundOfType.Floor)
                    If Not IsSignalTriggered(_fractalHighPayload(runningPayload.Value.PayloadDate) + buffer, Trade.TradeExecutionDirection.Buy, runningPayload.Value.PayloadDate, currentCandle.PayloadDate) Then
                        ret = True
                    End If
                    Exit For
                End If
            End If
        Next
        Return ret
    End Function

    Private Function IsFractalLowSatisfied(ByVal currentCandle As Payload) As Boolean
        Dim ret As Boolean = False
        For Each runningPayload In _signalPayload.OrderByDescending(Function(x)
                                                                        Return x.Key
                                                                    End Function)
            If runningPayload.Key < currentCandle.PayloadDate AndAlso runningPayload.Key.Date = _tradingDate.Date Then
                If _fractalLowPayload(runningPayload.Value.PayloadDate) > _fractalLowPayload(runningPayload.Value.PreviousCandlePayload.PayloadDate) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(_fractalLowPayload(runningPayload.Value.PayloadDate), RoundOfType.Floor)
                    If Not IsSignalTriggered(_fractalLowPayload(runningPayload.Value.PayloadDate) - buffer, Trade.TradeExecutionDirection.Sell, runningPayload.Value.PayloadDate, currentCandle.PayloadDate) Then
                        ret = True
                    End If
                    Exit For
                End If
            End If
        Next
        Return ret
    End Function
End Class
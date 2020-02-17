Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class FractalDipStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _fractalHighPayload As Dictionary(Of Date, Decimal)
    Private _fractalLowPayload As Dictionary(Of Date, Decimal)
    Private _atrPayload As Dictionary(Of Date, Decimal)

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
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
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
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If lastExecutedTrade Is Nothing Then
                    signalCandle = signalCandleSatisfied.Item4
                ElseIf lastExecutedTrade.SignalCandle.PayloadDate <> signalCandleSatisfied.Item4.PayloadDate Then
                    signalCandle = signalCandleSatisfied.Item4
                End If
                If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                    Dim atr As Decimal = ConvertFloorCeling(_atrPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim minATR As Decimal = signalCandleSatisfied.Item2 * 0.1 / 100
                    If atr > minATR Then
                        If signalCandleSatisfied.Item3 = Trade.TradeExecutionDirection.Buy Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                            Dim triggerPrice As Decimal = signalCandleSatisfied.Item2 + buffer
                            Dim slPoint As Decimal = atr + buffer
                            Dim stoplossPrice As Decimal = ConvertFloorCeling(triggerPrice - slPoint, _parentStrategy.TickSize, RoundOfType.Floor)
                            Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, triggerPrice, stoplossPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, _parentStrategy.StockType)
                            Dim targetPrice As Decimal = triggerPrice + (slPoint * _userInputs.TargetMultiplier)

                            parameter = New PlaceOrderParameters With {
                                        .EntryPrice = triggerPrice,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = quantity,
                                        .Stoploss = stoplossPrice,
                                        .Target = targetPrice,
                                        .Buffer = buffer,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.SL,
                                        .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                    }
                        ElseIf signalCandleSatisfied.Item3 = Trade.TradeExecutionDirection.Sell Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                            Dim triggerPrice As Decimal = signalCandleSatisfied.Item2 - buffer
                            Dim slPoint As Decimal = atr + buffer
                            Dim stoplossPrice As Decimal = ConvertFloorCeling(triggerPrice + slPoint, _parentStrategy.TickSize, RoundOfType.Floor)
                            Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoplossPrice, triggerPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, _parentStrategy.StockType)
                            Dim targetPrice As Decimal = triggerPrice - (slPoint * _userInputs.TargetMultiplier)

                            parameter = New PlaceOrderParameters With {
                                        .EntryPrice = triggerPrice,
                                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                        .Quantity = quantity,
                                        .Stoploss = stoplossPrice,
                                        .Target = targetPrice,
                                        .Buffer = buffer,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.SL,
                                        .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                    }
                        End If
                    Else
                        Console.WriteLine(String.Format("Trade neglected because of low atr: Time:{0}, ATR:{1}, Min required ATR:{2}, Instrument:{3}", signalCandle.PayloadDate, atr, minATR, _tradingSymbol))
                    End If
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim currentFractal As Decimal = _fractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                Dim signalFractal As Decimal = _fractalHighPayload(signalCandle.PayloadDate)
                If currentFractal <> signalFractal Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim currentFractal As Decimal = _fractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                Dim signalFractal As Decimal = _fractalLowPayload(signalCandle.PayloadDate)
                If currentFractal <> signalFractal Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
                Dim lastFractalHighPayload As Payload = Nothing
                Dim lastFractalLowPayload As Payload = Nothing
                Dim eligibleToTakeTrade As Boolean = False
                For Each runningPayload In _signalPayload.Keys
                    _cts.Token.ThrowIfCancellationRequested()
                    If runningPayload.Date = _tradingDate.Date AndAlso runningPayload <= candle.PayloadDate Then
                        If lastFractalLowPayload IsNot Nothing AndAlso _fractalHighPayload(runningPayload) < _fractalLowPayload(lastFractalLowPayload.PayloadDate) Then
                            eligibleToTakeTrade = True
                        ElseIf lastFractalHighPayload IsNot Nothing AndAlso _fractalLowPayload(runningPayload) > _fractalLowPayload(lastFractalHighPayload.PayloadDate) Then
                            eligibleToTakeTrade = True
                        End If
                        If lastFractalHighPayload Is Nothing Then
                            lastFractalHighPayload = _signalPayload(runningPayload)
                        Else
                            If _fractalHighPayload(runningPayload) <> _fractalHighPayload(_signalPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                lastFractalHighPayload = _signalPayload(runningPayload).PreviousCandlePayload
                            End If
                        End If
                        If lastFractalLowPayload Is Nothing Then
                            lastFractalLowPayload = _signalPayload(runningPayload)
                        Else
                            If _fractalLowPayload(runningPayload) <> _fractalLowPayload(_signalPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                lastFractalLowPayload = _signalPayload(runningPayload).PreviousCandlePayload
                            End If
                        End If
                    End If
                Next
                If eligibleToTakeTrade Then
                    If lastFractalHighPayload IsNot Nothing AndAlso lastFractalLowPayload IsNot Nothing Then
                        If lastFractalHighPayload.PayloadDate > lastFractalLowPayload.PayloadDate Then
                            ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, _fractalHighPayload(lastFractalHighPayload.PayloadDate), Trade.TradeExecutionDirection.Buy, lastFractalHighPayload)
                        Else
                            ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, _fractalLowPayload(lastFractalLowPayload.PayloadDate), Trade.TradeExecutionDirection.Sell, lastFractalLowPayload)
                        End If
                    ElseIf lastFractalHighPayload IsNot Nothing Then
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, _fractalHighPayload(lastFractalHighPayload.PayloadDate), Trade.TradeExecutionDirection.Buy, lastFractalHighPayload)
                    ElseIf lastFractalLowPayload IsNot Nothing Then
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, _fractalLowPayload(lastFractalLowPayload.PayloadDate), Trade.TradeExecutionDirection.Sell, lastFractalLowPayload)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

End Class
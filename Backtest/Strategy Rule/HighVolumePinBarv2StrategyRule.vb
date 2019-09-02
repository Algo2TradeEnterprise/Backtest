Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HighVolumePinBarv2StrategyRule
    Inherits StrategyRule

    Private _ATRPayload As Dictionary(Of Date, Decimal) = Nothing
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _ATRPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TradeType.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TradeType.MIS) AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > Math.Abs(_parentStrategy.OverAllLossPerDay) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime Then
            Dim signalCandle As Payload = Nothing
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TradeType.MIS)
            Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                If lastExecutedTrade Is Nothing Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                Else
                    Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedTrade.ExitTime, _signalPayload))
                    If exitCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                        signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                    ElseIf exitCandle.PayloadDate = currentMinuteCandlePayload.PayloadDate AndAlso _parentStrategy.RuleSupporting1 AndAlso
                        lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss AndAlso lastExecutedTrade.PLPoint < 0 Then
                        signalCandle = lastExecutedTrade.SignalCandle
                        signalCandleSatisfied = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, lastExecutedTrade.EntryDirection)
                    End If
                End If
            ElseIf _parentStrategy.RuleSupporting1 AndAlso lastExecutedTrade IsNot Nothing AndAlso
                lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss AndAlso lastExecutedTrade.PLPoint < 0 Then
                signalCandle = lastExecutedTrade.SignalCandle
                signalCandleSatisfied = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, lastExecutedTrade.EntryDirection)
            End If
            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim targetPoint As Decimal = ConvertFloorCeling(_ATRPayload(signalCandle.PayloadDate) * _parentStrategy.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                Dim potentialSLPoint As Decimal = ConvertFloorCeling(_ATRPayload(signalCandle.PayloadDate) * _parentStrategy.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)

                If signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandle.High + buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = _lotSize,
                        .Stoploss = Math.Max(signalCandle.Low - buffer, .EntryPrice - potentialSLPoint),
                        .Target = .EntryPrice + targetPoint,
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString
                    }
                ElseIf signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandle.Low - buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                        .Quantity = _lotSize,
                        .Stoploss = Math.Min(signalCandle.High + buffer, .EntryPrice + potentialSLPoint),
                        .Target = .EntryPrice - targetPoint,
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString
                    }
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TradeType.MIS)
                If lastExecutedTrade Is Nothing OrElse lastExecutedTrade.SignalCandle.PayloadDate <> signalCandle.PayloadDate Then
                    If signalCandle.PayloadDate = currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PayloadDate Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                ElseIf _parentStrategy.RuleSupporting1 Then
                    Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedTrade.ExitTime, _signalPayload))
                    If exitCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                        Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
                        If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                            If signalCandle.PayloadDate <> currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
                                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If _parentStrategy.ModifyStoploss AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim breakoutCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTrade.EntryTime, _signalPayload))
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If breakoutCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                signalCandle = breakoutCandle
            End If
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTrade.EntryPrice, RoundOfType.Floor)
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                triggerPrice = signalCandle.Low - buffer
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                triggerPrice = signalCandle.High + buffer
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Move for breakout candle: {0}. Time:{1}", triggerPrice, currentTick.PayloadDate))
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function IsSignalCandle(ByVal candle As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection) = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(False, Trade.TradeExecutionDirection.None)
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If candle.Volume > candle.PreviousCandlePayload.Volume Then
                If (candle.CandleWicks.Top / candle.CandleRange) * 100 >= 50 Then
                    If candle.High > candle.PreviousCandlePayload.High AndAlso
                        candle.Low > candle.PreviousCandlePayload.Low Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Sell)
                    End If
                ElseIf (candle.CandleWicks.Bottom / candle.CandleRange) * 100 >= 50 Then
                    If candle.Low < candle.PreviousCandlePayload.Low AndAlso
                        candle.High < candle.PreviousCandlePayload.High Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Buy)
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class

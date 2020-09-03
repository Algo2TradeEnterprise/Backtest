Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class InsideBarStrategyRule
    Inherits StrategyRule

    Private _10MinPayloads As Dictionary(Of Date, Payload) = Nothing
    Private _15MinPayloads As Dictionary(Of Date, Payload) = Nothing

    Private _pivotPointsPayload As Dictionary(Of Date, PivotPoints) = Nothing

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Dim exchangeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day,
                                                 _parentStrategy.ExchangeStartTime.Hours, _parentStrategy.ExchangeStartTime.Minutes, _parentStrategy.ExchangeStartTime.Seconds)

        _10MinPayloads = Common.ConvertPayloadsToXMinutes(_inputPayload, 10, exchangeStartTime)
        _15MinPayloads = Common.ConvertPayloadsToXMinutes(_inputPayload, 15, exchangeStartTime)
        Indicator.Pivots.CalculatePivots(_signalPayload, _pivotPointsPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Decimal, String) = GetEntrySignal(currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandle, Trade.TypeOfTrade.MIS)
                If lastExecutedOrder IsNot Nothing Then
                    If lastExecutedOrder.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                        signalCandle = signal.Item2
                    End If
                Else
                    signalCandle = signal.Item2
                End If
            End If
            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                Dim entryPrice As Decimal = currentTick.Open
                Dim stoploss As Decimal = signal.Item3 - buffer
                Dim quantity As Integer = Me.LotSize
                Dim targetPoint As Decimal = 5

                parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = .EntryPrice + targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = signal.Item4
                            }

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
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
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

    Private Function GetEntrySignal(ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Decimal, String)
        Dim ret As Tuple(Of Boolean, Payload, Decimal, String) = Nothing
        Dim current1MinCandle As Payload = _inputPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _inputPayload, 1))
        Dim current5MinCandle As Payload = Nothing
        If _signalPayload.ContainsKey(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload, 5)) Then
            current5MinCandle = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload, 5))
        End If
        Dim current10MinCandle As Payload = Nothing
        If _10MinPayloads.ContainsKey(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _10MinPayloads, 10)) Then
            current10MinCandle = _10MinPayloads(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _10MinPayloads, 10))
        End If
        Dim current15MinCandle As Payload = Nothing
        If _15MinPayloads.ContainsKey(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _15MinPayloads, 15)) Then
            current15MinCandle = _15MinPayloads(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _15MinPayloads, 15))
        End If
        If current5MinCandle IsNot Nothing AndAlso current5MinCandle.PreviousCandlePayload IsNot Nothing AndAlso
            current5MinCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            current5MinCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
            Dim pivot As PivotPoints = _pivotPointsPayload(current5MinCandle.PreviousCandlePayload.PayloadDate)
            Dim stoploss As Decimal = current5MinCandle.PreviousCandlePayload.PreviousCandlePayload.Low
            If IsSignalValid(current5MinCandle, current1MinCandle, currentTick, pivot) Then
                ret = New Tuple(Of Boolean, Payload, Decimal, String)(True, current5MinCandle.PreviousCandlePayload, stoploss, "Signal at 5 min candle")
            ElseIf IsSignalValid(current10MinCandle, current1MinCandle, currentTick, pivot) Then
                ret = New Tuple(Of Boolean, Payload, Decimal, String)(True, current10MinCandle.PreviousCandlePayload, stoploss, "Signal at 10 min candle")
            ElseIf IsSignalValid(current15MinCandle, current1MinCandle, currentTick, pivot) Then
                ret = New Tuple(Of Boolean, Payload, Decimal, String)(True, current15MinCandle.PreviousCandlePayload, stoploss, "Signal at 15 min candle")
            End If
        End If
        Return ret
    End Function

    Private Function IsSignalValid(ByVal currentXMinuteCandle As Payload, ByVal current1MinuteCandle As Payload, ByVal currentTick As Payload, ByVal pivot As PivotPoints) As Boolean
        Dim ret As Boolean = False
        If currentXMinuteCandle IsNot Nothing AndAlso currentXMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
            If currentXMinuteCandle.PreviousCandlePayload.High < currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.High AndAlso
                currentXMinuteCandle.PreviousCandlePayload.Low > currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.Low Then
                If currentXMinuteCandle.PreviousCandlePayload.Close > currentXMinuteCandle.PreviousCandlePayload.Open Then
                    If currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.High - currentXMinuteCandle.PreviousCandlePayload.Close >= 2 AndAlso
                        currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.High - currentXMinuteCandle.PreviousCandlePayload.Close < 4 Then
                        If currentXMinuteCandle.PreviousCandlePayload.Close >= pivot.Resistance1 Then
                            If currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.Close - currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.Low >= 3 Then
                                If currentXMinuteCandle.Open >= currentXMinuteCandle.PreviousCandlePayload.Close Then
                                    Dim tickTime As Date = New Date(currentXMinuteCandle.PayloadDate.Year, currentXMinuteCandle.PayloadDate.Month, currentXMinuteCandle.PayloadDate.Day,
                                                                    currentXMinuteCandle.PayloadDate.Hour, currentXMinuteCandle.PayloadDate.Minute, 10)
                                    If currentTick.PayloadDate >= tickTime AndAlso current1MinuteCandle.Volume > 0 AndAlso
                                        currentTick.Close >= currentXMinuteCandle.PreviousCandlePayload.Close Then
                                        ret = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class
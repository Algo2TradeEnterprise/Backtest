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
            Dim signal As Tuple(Of Boolean, Payload, Decimal, String, Trade.TradeExecutionDirection) = GetEntrySignal(currentTick)
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
                If signal.Item5 = Trade.TradeExecutionDirection.Buy Then
                    Dim slPoint As Decimal = Math.Min((entryPrice - (signal.Item3 - buffer)), 10)
                    Dim quantity As Integer = Me.LotSize
                    Dim targetPoint As Decimal = 100000000

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item4
                                }
                ElseIf signal.Item5 = Trade.TradeExecutionDirection.Sell Then
                    Dim slPoint As Decimal = Math.Min(ConvertFloorCeling((entryPrice * 1.005) - entryPrice, _parentStrategy.TickSize, RoundOfType.Floor), 10)
                    Dim quantity As Integer = Me.LotSize
                    Dim targetPoint As Decimal = 100000000

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item4
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
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim remark As String = ""
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim plPoint As Decimal = currentTick.Open - currentTrade.EntryPrice
                Dim plPer As Decimal = (plPoint / currentTrade.EntryPrice) * 100
                If plPer >= 0.5 AndAlso plPer <= 10 Then
                    Dim gainPer As Decimal = Math.Floor((plPer * 100) / 10)
                    If gainPer Mod 2 = 0 Then
                        gainPer = (gainPer - 1) / 10
                    Else
                        gainPer = gainPer / 10
                    End If
                    If plPer > gainPer Then
                        Dim slPer As Decimal = gainPer
                        Dim potentialStoploss As Decimal = currentTrade.EntryPrice + ConvertFloorCeling(currentTrade.EntryPrice * slPer / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                        If potentialStoploss > currentTrade.PotentialStopLoss Then
                            triggerPrice = potentialStoploss
                            remark = String.Format("Moved at {0} for gain {1}%", currentTick.PayloadDate.ToString("HH:mm:ss"), Math.Round(gainPer, 2))
                        End If
                    End If
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim plPoint As Decimal = currentTrade.EntryPrice - currentTick.Open
                Dim plPer As Decimal = (plPoint / currentTrade.EntryPrice) * 100
                If plPer >= 0.5 AndAlso plPer <= 10 Then
                    Dim gainPer As Decimal = Math.Floor((plPer * 100) / 10)
                    If gainPer Mod 2 = 0 Then
                        gainPer = (gainPer - 1) / 10
                    Else
                        gainPer = gainPer / 10
                    End If
                    If plPer > gainPer Then
                        Dim slPer As Decimal = gainPer
                        Dim potentialStoploss As Decimal = currentTrade.EntryPrice - ConvertFloorCeling(currentTrade.EntryPrice * slPer / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                        If potentialStoploss < currentTrade.PotentialStopLoss Then
                            triggerPrice = potentialStoploss
                            remark = String.Format("Moved at {0} for gain {1}%", currentTick.PayloadDate.ToString("HH:mm:ss"), Math.Round(gainPer, 2))
                        End If
                    End If
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, remark)
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

    Private Function GetEntrySignal(ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Decimal, String, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Payload, Decimal, String, Trade.TradeExecutionDirection) = Nothing
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
            Dim direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
            If IsSignalValid(current5MinCandle, current1MinCandle, currentTick, pivot, direction) Then
                Dim stoploss As Decimal = current5MinCandle.PreviousCandlePayload.PreviousCandlePayload.Low
                ret = New Tuple(Of Boolean, Payload, Decimal, String, Trade.TradeExecutionDirection)(True, current5MinCandle.PreviousCandlePayload, stoploss, "Signal at 5 min candle", direction)
            ElseIf IsSignalValid(current10MinCandle, current1MinCandle, currentTick, pivot, direction) Then
                Dim stoploss As Decimal = current10MinCandle.PreviousCandlePayload.PreviousCandlePayload.Low
                ret = New Tuple(Of Boolean, Payload, Decimal, String, Trade.TradeExecutionDirection)(True, current10MinCandle.PreviousCandlePayload, stoploss, "Signal at 10 min candle", direction)
            ElseIf IsSignalValid(current15MinCandle, current1MinCandle, currentTick, pivot, direction) Then
                Dim stoploss As Decimal = current15MinCandle.PreviousCandlePayload.PreviousCandlePayload.Low
                ret = New Tuple(Of Boolean, Payload, Decimal, String, Trade.TradeExecutionDirection)(True, current15MinCandle.PreviousCandlePayload, stoploss, "Signal at 15 min candle", direction)
            End If
        End If
        Return ret
    End Function

    Private Function IsSignalValid(ByVal currentXMinuteCandle As Payload, ByVal current1MinuteCandle As Payload,
                                   ByVal currentTick As Payload, ByVal pivot As PivotPoints, ByRef direction As Trade.TradeExecutionDirection) As Boolean
        Dim ret As Boolean = False
        If currentXMinuteCandle IsNot Nothing AndAlso currentXMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
            If currentXMinuteCandle.PreviousCandlePayload.High < currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.High AndAlso
                currentXMinuteCandle.PreviousCandlePayload.Low > currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.Low Then
                If currentXMinuteCandle.PreviousCandlePayload.Close > currentXMinuteCandle.PreviousCandlePayload.Open Then
                    If currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.High - currentXMinuteCandle.PreviousCandlePayload.Close >= 2 AndAlso
                        currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.High - currentXMinuteCandle.PreviousCandlePayload.Close < 4 Then
                        If currentXMinuteCandle.PreviousCandlePayload.Close >= pivot.Pivot Then
                            If currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.Close - currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.Low >= 3 Then
                                If currentXMinuteCandle.Open >= currentXMinuteCandle.PreviousCandlePayload.Close Then
                                    Dim tickTime As Date = New Date(currentXMinuteCandle.PayloadDate.Year, currentXMinuteCandle.PayloadDate.Month, currentXMinuteCandle.PayloadDate.Day,
                                                                    currentXMinuteCandle.PayloadDate.Hour, currentXMinuteCandle.PayloadDate.Minute, 10)
                                    If currentTick.PayloadDate >= tickTime AndAlso current1MinuteCandle.Volume > 0 AndAlso
                                        currentTick.Close >= currentXMinuteCandle.PreviousCandlePayload.Close Then
                                        'ret=True
                                        'New Condition
                                        'If currentXMinuteCandle.Open < Math.Min(pivot.Resistance1, Math.Min(pivot.Resistance2, pivot.Resistance3)) AndAlso
                                        '    Math.Min(pivot.Resistance1, Math.Min(pivot.Resistance2, pivot.Resistance3)) - currentXMinuteCandle.Open >= 8 Then
                                        Dim preCndlHighBelowPivot As Boolean = False
                                        If currentXMinuteCandle.PreviousCandlePayload.Open < pivot.Pivot Then
                                            If currentXMinuteCandle.PreviousCandlePayload.High < pivot.Pivot Then
                                                preCndlHighBelowPivot = True
                                            End If
                                        ElseIf currentXMinuteCandle.PreviousCandlePayload.Open < pivot.Resistance1 Then
                                            If currentXMinuteCandle.PreviousCandlePayload.High < pivot.Resistance1 Then
                                                preCndlHighBelowPivot = True
                                            End If
                                        ElseIf currentXMinuteCandle.PreviousCandlePayload.Open < pivot.Resistance2 Then
                                            If currentXMinuteCandle.PreviousCandlePayload.High < pivot.Resistance2 Then
                                                preCndlHighBelowPivot = True
                                            End If
                                        ElseIf currentXMinuteCandle.PreviousCandlePayload.Open < pivot.Resistance3 Then
                                            If currentXMinuteCandle.PreviousCandlePayload.High < pivot.Resistance3 Then
                                                preCndlHighBelowPivot = True
                                            End If
                                        End If
                                        'If preCndlHighBelowPivot Then
                                        Dim prePreCndlHighBelowPivot As Boolean = False
                                        If currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.Open < pivot.Pivot Then
                                            If currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.High < pivot.Pivot Then
                                                prePreCndlHighBelowPivot = True
                                            End If
                                        ElseIf currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.Open < pivot.Resistance1 Then
                                            If currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.High < pivot.Resistance1 Then
                                                prePreCndlHighBelowPivot = True
                                            End If
                                        ElseIf currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.Open < pivot.Resistance2 Then
                                            If currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.High < pivot.Resistance2 Then
                                                prePreCndlHighBelowPivot = True
                                            End If
                                        ElseIf currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.Open < pivot.Resistance3 Then
                                            If currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.High < pivot.Resistance3 Then
                                                prePreCndlHighBelowPivot = True
                                            End If
                                        End If
                                        If preCndlHighBelowPivot OrElse prePreCndlHighBelowPivot Then
                                            'If currentXMinuteCandle.PreviousCandlePayload.High < pivot.Resistance3 AndAlso
                                            '    currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.High < pivot.Resistance3 Then
                                            If currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Green Then
                                                direction = Trade.TradeExecutionDirection.Buy
                                            Else
                                                If currentXMinuteCandle.PreviousCandlePayload.CandleBody > currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.CandleBody * 0.6 Then
                                                    direction = Trade.TradeExecutionDirection.Sell
                                                Else
                                                    direction = Trade.TradeExecutionDirection.Buy
                                                End If
                                            End If
                                            If direction <> Trade.TradeExecutionDirection.None Then
                                                If currentXMinuteCandle.PreviousCandlePayload.CandleBody > 1.15 Then
                                                    '(20) no checked as it is same with (18)
                                                    If Not (currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Green AndAlso
                                                        currentXMinuteCandle.PreviousCandlePayload.CandleBody > currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.CandleBody) Then
                                                        If currentXMinuteCandle.PreviousCandlePayload.Open >= currentXMinuteCandle.PreviousCandlePayload.PreviousCandlePayload.Close Then
                                                            ret = True
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            'End If
                                        End If
                                        'End If
                                        'End If
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
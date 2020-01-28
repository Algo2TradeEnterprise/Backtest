Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PositionalHourlyStrategyRule
    Inherits StrategyRule

    Private _smaPayload As Dictionary(Of Date, Decimal)

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

        Indicator.SMA.CalculateSMA(200, Payload.PayloadFields.Close, _signalPayload, _smaPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) AndAlso Not _parentStrategy.IsTradeActive(currentTick, _parentStrategy.TradeType) Then
            Dim signalCandle As Payload = Nothing

            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, _parentStrategy.TradeType)
            Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Payload) = GetSignalForEntry(currentTick)
            If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                signalCandle = signalReceivedForEntry.Item3
                If signalCandle IsNot Nothing Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalReceivedForEntry.Item2, RoundOfType.Floor)
                    Dim tradeNumber As Integer = 1
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromInvestment(Me.LotSize, 100000, signalReceivedForEntry.Item2, Trade.TypeOfStock.Cash, True)
                    Dim slPoint As Decimal = signalReceivedForEntry.Item3.CandleRange + 2 * buffer
                    Dim lowestSLPoint As Decimal = slPoint
                    If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                        tradeNumber = lastExecutedTrade.Supporting2 + 1
                        'quantity = Math.Ceiling(lastExecutedTrade.Quantity / 2)
                        quantity = lastExecutedTrade.Quantity
                        lowestSLPoint = Math.Min(lowestSLPoint, CDec(lastExecutedTrade.Supporting1))
                    End If
                    Dim targetPoint As Decimal = lowestSLPoint * tradeNumber

                    Dim takeTrade As Boolean = False
                    If currentMinuteCandlePayload.Open < signalReceivedForEntry.Item2 Then
                        takeTrade = True
                    Else
                        Dim currentOneMinutePayload As Payload = _inputPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _inputPayload))
                        Dim payloadsToCheck As IEnumerable(Of Payload) = _inputPayload.Values.Where(Function(x)
                                                                                                        Return x.PayloadDate >= currentMinuteCandlePayload.PayloadDate AndAlso
                                                                                                        x.PayloadDate < currentOneMinutePayload.PayloadDate AndAlso
                                                                                                        x.PayloadDate >= currentOneMinutePayload.PayloadDate.AddMinutes(-5)
                                                                                                    End Function)
                        If payloadsToCheck IsNot Nothing AndAlso payloadsToCheck.Count > 0 Then
                            Dim high As Decimal = payloadsToCheck.Max(Function(x)
                                                                          Return x.High
                                                                      End Function)

                            If high < signalReceivedForEntry.Item2 Then
                                takeTrade = True
                            End If
                        End If
                    End If
                    If takeTrade AndAlso (tradeNumber > 1 OrElse signalCandle.Close > _smaPayload(signalCandle.PayloadDate)) Then
                        parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalReceivedForEntry.Item2,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - slPoint,
                                .Target = .EntryPrice + targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = lowestSLPoint,
                                .Supporting2 = tradeNumber,
                                .Supporting3 = signalReceivedForEntry.Item3.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss")
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Payload) = GetSignalForEntry(currentTick)
            If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signalReceivedForEntry.Item3.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            Else
                Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                If currentMinuteCandlePayload.Close < _smaPayload(currentMinuteCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
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

#Region "Entry Rule"
    Private Function GetSignalForEntry(ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Payload) = Nothing
        Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, _parentStrategy.TradeType)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim lowerHighCandle As Payload = GetPreviousLowerHigh(currentMinuteCandlePayload)
        If lowerHighCandle IsNot Nothing Then
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(lowerHighCandle.High, RoundOfType.Floor)
            Dim entryPrice As Decimal = lowerHighCandle.High + buffer
            If lastExecutedTrade Is Nothing Then
                ret = New Tuple(Of Boolean, Decimal, Payload)(True, entryPrice, lowerHighCandle)
            Else
                Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedTrade.ExitTime, _signalPayload))
                If lowerHighCandle.PayloadDate >= exitCandle.PayloadDate Then
                    ret = New Tuple(Of Boolean, Decimal, Payload)(True, entryPrice, lowerHighCandle)
                End If
            End If
        Else
            If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                Dim entryCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedTrade.EntryTime, _signalPayload))
                Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedTrade.ExitTime, _signalPayload))
                If entryCandle.Open > entryCandle.PreviousCandlePayload.Close Then
                    Dim otherLowerHighFound As Boolean = False
                    For Each runningPayload In _signalPayload.Values
                        If runningPayload.PayloadDate >= entryCandle.PayloadDate AndAlso
                            runningPayload.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                            If runningPayload.High < runningPayload.PreviousCandlePayload.High Then
                                otherLowerHighFound = True
                            End If
                        End If
                    Next
                    If Not otherLowerHighFound Then
                        If currentMinuteCandlePayload.PayloadDate > exitCandle.PayloadDate Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(lastExecutedTrade.SignalCandle.High, RoundOfType.Floor)
                            Dim entryPrice As Decimal = lastExecutedTrade.SignalCandle.High + buffer
                            ret = New Tuple(Of Boolean, Decimal, Payload)(True, entryPrice, lastExecutedTrade.SignalCandle)
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetPreviousLowerHigh(ByVal currentCandle As Payload) As Payload
        Dim ret As Payload = currentCandle.PreviousCandlePayload
        While True
            If ret IsNot Nothing AndAlso ret.PreviousCandlePayload IsNot Nothing Then
                If ret.High < ret.PreviousCandlePayload.High Then
                    Exit While
                ElseIf ret.High = ret.PreviousCandlePayload.High Then
                    ret = ret.PreviousCandlePayload
                Else
                    ret = Nothing
                    Exit While
                End If
            Else
                Exit While
            End If
        End While
        Return ret
    End Function
#End Region

End Class

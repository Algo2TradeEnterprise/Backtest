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
                If lastExecutedTrade Is Nothing OrElse lastExecutedTrade.SignalCandle.PayloadDate <> signalReceivedForEntry.Item3.PayloadDate Then
                    signalCandle = signalReceivedForEntry.Item3
                End If

                If signalCandle IsNot Nothing AndAlso currentMinuteCandlePayload.Open < signalReceivedForEntry.Item2 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalReceivedForEntry.Item2, RoundOfType.Floor)
                    Dim quantity As Integer = 1
                    Dim slPoint As Decimal = signalReceivedForEntry.Item3.CandleRange + 2 * buffer
                    Dim lowestSLPoint As Decimal = slPoint
                    If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                        quantity = lastExecutedTrade.Quantity + 1
                        lowestSLPoint = Math.Max(lowestSLPoint, CDec(lastExecutedTrade.Supporting1))
                    End If
                    Dim targetPoint As Decimal = lowestSLPoint

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
                        .Supporting2 = signalReceivedForEntry.Item3.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss")
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim lowerHighCandle As Payload = GetPreviousLowerHigh(currentMinuteCandlePayload)
        If lowerHighCandle IsNot Nothing AndAlso lowerHighCandle.Close > _smaPayload(lowerHighCandle.PayloadDate) Then
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(lowerHighCandle.High, RoundOfType.Floor)
            Dim entryPrice As Decimal = lowerHighCandle.High + buffer
            'If currentTick.Open >= entryPrice Then
            'If currentMinuteCandlePayload.Open <= lowerHighCandle.High Then
            ret = New Tuple(Of Boolean, Decimal, Payload)(True, entryPrice, lowerHighCandle)
            'Else
            '    Dim totalPayloadsBelowEntryPrice As IEnumerable(Of Payload) = _inputPayload.Values.Where(Function(x)
            '                                                                                                 If x.PayloadDate >= currentMinuteCandlePayload.PayloadDate AndAlso
            '                                                                                                    x.PayloadDate < currentTick.PayloadDate AndAlso
            '                                                                                                    x.PayloadDate >= currentTick.PayloadDate.AddMinutes(-5) AndAlso
            '                                                                                                    x.High <= lowerHighCandle.High Then
            '                                                                                                     Return True
            '                                                                                                 Else
            '                                                                                                     Return Nothing
            '                                                                                                 End If
            '                                                                                             End Function)
            '    If totalPayloadsBelowEntryPrice IsNot Nothing AndAlso totalPayloadsBelowEntryPrice.Count >= 5 Then
            '        ret = New Tuple(Of Boolean, Decimal, Payload)(True, lowerHighCandle.High, lowerHighCandle)
            '    End If
            'End If
            'End If
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

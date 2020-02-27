Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HKPositionalHourlyStrategyRule2
    Inherits StrategyRule

    Private _hkPayload As Dictionary(Of Date, Payload)

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

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
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

            Dim signal As Tuple(Of Boolean, Decimal, Payload) = GetSignalForEntry(currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item3
                If signalCandle IsNot Nothing Then
                    'Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                    Dim buffer As Decimal = 1
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = currentTick.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = Me.LotSize,
                                .Stoploss = ConvertFloorCeling(signalCandle.Open - buffer, _parentStrategy.TickSize, RoundOfType.Floor),
                                .Target = .EntryPrice + 1000000,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market
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
        'If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
        '    Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Payload) = GetSignalForEntry(currentTick)
        '    If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
        '        If currentTrade.SignalCandle.PayloadDate <> signalReceivedForEntry.Item3.PayloadDate Then
        '            ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
        '        End If
        '    Else
        '        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        '        If currentMinuteCandlePayload.Close < _smaPayload(currentMinuteCandlePayload.PayloadDate) Then
        '            ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
        '        End If
        '    End If
        'End If
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim currentHKPayload As Payload = _hkPayload(currentMinuteCandlePayload.PayloadDate)
            If currentTrade.PotentialStopLoss <> currentHKPayload.Open - currentTrade.StoplossBuffer Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, currentHKPayload.Open - currentTrade.StoplossBuffer, String.Format("Modified at {0}",
                                                                                                        currentTick.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss")))
            End If
        End If
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
        Dim currentHKPayload As Payload = _hkPayload(currentMinuteCandlePayload.PayloadDate)
        If currentTick.PayloadDate >= currentHKPayload.PayloadDate.AddMinutes(59) Then
            If currentHKPayload.Open = currentHKPayload.Low Then
                ret = New Tuple(Of Boolean, Decimal, Payload)(True, currentTick.Open, currentHKPayload)
            End If
        End If
        Return ret
    End Function
#End Region

End Class

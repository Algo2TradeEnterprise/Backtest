Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class SwingStrategyRule
    Inherits StrategyRule

    Private _swingHighPayload As Dictionary(Of Date, Decimal)
    Private _swingLowPayload As Dictionary(Of Date, Decimal)
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

        Indicator.SwingHighLow.CalculateSwingHighLow(_signalPayload, True, _swingHighPayload, _swingLowPayload)
        Indicator.SMA.CalculateSMA(200, Payload.PayloadFields.Close, _signalPayload, _smaPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, _parentStrategy.TradeType) Then
            Dim signal As Tuple(Of Boolean, Decimal) = GetSignalForEntry(currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                parameter = New PlaceOrderParameters With {
                    .EntryPrice = signal.Item2,
                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                    .Quantity = Me.LotSize,
                    .Stoploss = .EntryPrice - 1000000,
                    .Target = .EntryPrice + 1000000,
                    .Buffer = 0,
                    .SignalCandle = currentMinuteCandlePayload,
                    .OrderType = Trade.TypeOfOrder.Market
                }
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 AndAlso _signalPayload.ContainsKey(currentTick.PayloadDate.Date) Then
                Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
                If currentDayPayload.Low <= currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, currentTrade.PotentialStopLoss, "Stoploss")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 AndAlso _signalPayload.ContainsKey(currentTick.PayloadDate.Date) Then
                Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
                Dim sma As Decimal = _smaPayload(currentDayPayload.PayloadDate)
                Dim swingLow As Decimal = _swingLowPayload(currentDayPayload.PayloadDate)
                If swingLow < sma Then
                    If swingLow > currentTrade.PotentialStopLoss Then
                        ret = New Tuple(Of Boolean, Decimal, String)(True, swingLow, String.Format("Modified at {0}", currentDayPayload.PayloadDate.ToString("dd-MM-yyyy")))
                    End If
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

    Private Function GetSignalForEntry(ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal)
        Dim ret As Tuple(Of Boolean, Decimal) = Nothing
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 AndAlso _signalPayload.ContainsKey(currentTick.PayloadDate.Date) Then
            Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
            Dim sma As Decimal = _smaPayload(currentDayPayload.PayloadDate)
            If currentDayPayload.Close > sma AndAlso currentDayPayload.PreviousCandlePayload.Close < sma Then
                ret = New Tuple(Of Boolean, Decimal)(True, currentDayPayload.Close)
            End If
        End If
        Return ret
    End Function
End Class

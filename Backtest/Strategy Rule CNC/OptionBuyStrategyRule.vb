Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers
Imports Backtest.StrategyHelper

Public Class OptionBuyStrategyRule
    Inherits StrategyRule

    Public Sub New(ByVal canceller As CancellationTokenSource,
                    ByVal tradingDate As Date,
                    ByVal nextTradingDay As Date,
                    ByVal tradingSymbol As String,
                    ByVal lotSize As Integer,
                    ByVal parentStrategy As Strategy,
                    ByVal inputMinPayload As Dictionary(Of Date, Payload),
                    ByVal inputEODPayload As Dictionary(Of Date, Payload))
        MyBase.New(canceller, tradingDate, nextTradingDay, tradingSymbol, lotSize, parentStrategy, inputMinPayload, inputEODPayload)
    End Sub

    Public Overrides Async Function UpdateCollectionsIfRequiredAsync(currentTickTime As Date) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Protected Overrides Function IsEntrySignalReceived(currentTickTime As Date) As Tuple(Of Boolean, Payload, Trade.TradeDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeDirection) = Nothing
        Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime, _ParentStrategy.SignalTimeFrame)
        If _SignalPayload IsNot Nothing AndAlso _SignalPayload.ContainsKey(currentMinute) Then
            ret = New Tuple(Of Boolean, Payload, Trade.TradeDirection)(True, _SignalPayload(currentMinute), Trade.TradeDirection.Buy)
        End If
        Return ret
    End Function
End Class
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PairStrategyRule
    Inherits StrategyRule

    Private ReadOnly _direction As Trade.TradeExecutionDirection
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal anotherPairInstrument As StrategyRule,
                   ByVal canceller As CancellationTokenSource,
                   ByVal direction As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, anotherPairInstrument, canceller)
        If direction > 1 Then
            _direction = Trade.TradeExecutionDirection.Buy
        Else
            _direction = Trade.TradeExecutionDirection.Sell
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If Me.ForceTakeTrade Then
            Dim quantity As Integer = _parentStrategy.CalculateQuantityFromInvestment(Me.LotSize, 10000, currentTick.Open, _parentStrategy.StockType, True)
            Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                        .EntryPrice = currentTick.Open,
                                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                        .Quantity = quantity,
                                                        .Stoploss = .EntryPrice - 100000,
                                                        .Target = .EntryPrice + 100000,
                                                        .Buffer = 0,
                                                        .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                                        .OrderType = Trade.TypeOfOrder.Market,
                                                        .Supporting1 = "Force entry"
                                                    }

            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        Else

        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function
End Class

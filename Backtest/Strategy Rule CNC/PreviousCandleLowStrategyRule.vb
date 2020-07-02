Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class PreviousCandleLowStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public InitialCapital As Integer
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private ReadOnly _userInputs As StrategyRuleEntities

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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload, True)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) Then
            Dim signal As Tuple(Of Boolean, Decimal, Integer, Integer, String) = GetSignalForEntry(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim entryPrice As Decimal = signal.Item2
                Dim quantity As Integer = signal.Item3
                Dim iteration As Integer = signal.Item4
                Dim remarks As String = signal.Item5

                parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - 1000000000,
                                .Target = .EntryPrice + 1000000000,
                                .Buffer = 0,
                                .SignalCandle = currentMinuteCandlePayload,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = iteration,
                                .Supporting2 = remarks
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

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
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

    Private Function GetSignalForEntry(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Integer, Integer, String)
        Dim ret As Tuple(Of Boolean, Decimal, Integer, Integer, String) = Nothing
        If candle.Low < candle.PreviousCandlePayload.Low Then
            If currentTick.Low < candle.PreviousCandlePayload.Low Then
                Dim lastTrade As Trade = GetLastOrder(candle)
                Dim atr As Decimal = _atrPayload(candle.PreviousCandlePayload.PayloadDate)
                Dim iteration As Integer = 1
                Dim quantity As Integer = GetQuantity(iteration, candle.Open)

                Dim entryPrice As Decimal = currentTick.Low
                If lastTrade Is Nothing Then
                    ret = New Tuple(Of Boolean, Decimal, Integer, Integer, String)(True, entryPrice, quantity, iteration, "First Trade")
                Else
                    If lastTrade.SignalCandle.PayloadDate <> candle.PayloadDate Then
                        If entryPrice < lastTrade.EntryPrice Then
                            If lastTrade.EntryPrice - entryPrice >= atr Then
                                iteration = Val(lastTrade.Supporting1) + 1
                                quantity = GetQuantity(iteration, entryPrice)
                                ret = New Tuple(Of Boolean, Decimal, Integer, Integer, String)(True, entryPrice, quantity, iteration, "Below last entry")
                            End If
                        Else
                            ret = New Tuple(Of Boolean, Decimal, Integer, Integer, String)(True, entryPrice, quantity, iteration, "(Reset) Above last entry")
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetQuantity(ByVal iterationNumber As Integer, ByVal price As Decimal) As Integer
        Dim capital As Decimal = _userInputs.InitialCapital * iterationNumber
        Return _parentStrategy.CalculateQuantityFromInvestment(1, capital, price, Trade.TypeOfStock.Cash, False)
    End Function

    Private Function GetLastOrder(ByVal currentPayload As Payload) As Trade
        Dim ret As Trade = Nothing
        If currentPayload IsNot Nothing Then
            ret = Me._parentStrategy.GetLastExecutedTradeOfTheStock(currentPayload, Me._parentStrategy.TradeType)
        End If
        Return ret
    End Function
End Class

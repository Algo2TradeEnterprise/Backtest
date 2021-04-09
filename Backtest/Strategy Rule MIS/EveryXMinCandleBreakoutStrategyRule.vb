Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class EveryXMinCandleBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public Capital As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso Me.EligibleToTakeTrade AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime Then
            Dim signalCandle As Payload = currentMinuteCandle.PreviousCandlePayload
            Dim lastTrade As Trade = _parentStrategy.GetLastTradeOfTheStock(currentMinuteCandle, Trade.TypeOfTrade.MIS)
            If lastTrade IsNot Nothing AndAlso lastTrade.SignalCandle.PayloadDate = signalCandle.PayloadDate Then
                signalCandle = Nothing
            End If
            If signalCandle IsNot Nothing Then
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromInvestment(Me.LotSize, _userInputs.Capital, signalCandle.Close, Trade.TypeOfStock.Cash, True)
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Close, RoundOfType.Celing)

                Dim buyPrice As Decimal = signalCandle.High + buffer
                Dim sellPrice As Decimal = signalCandle.Low - buffer
                Dim entryDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
                If currentMinuteCandle.High >= buyPrice AndAlso currentMinuteCandle.Low <= sellPrice Then
                    If currentMinuteCandle.CandleColor = Color.Green Then
                        entryDirection = Trade.TradeExecutionDirection.Sell
                    Else
                        entryDirection = Trade.TradeExecutionDirection.Buy
                    End If
                ElseIf currentMinuteCandle.High >= buyPrice Then
                    entryDirection = Trade.TradeExecutionDirection.Buy
                ElseIf currentMinuteCandle.Low <= sellPrice Then
                    entryDirection = Trade.TradeExecutionDirection.Sell
                End If

                If entryDirection = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                            .EntryPrice = buyPrice,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice - 1000000,
                            .Target = .EntryPrice + 1000000,
                            .Buffer = buffer,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.SL,
                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                        }
                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                ElseIf entryDirection = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                            .EntryPrice = sellPrice,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice + 1000000,
                            .Target = .EntryPrice - 1000000,
                            .Buffer = buffer,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.SL,
                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                        }
                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                End If
            End If
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
End Class
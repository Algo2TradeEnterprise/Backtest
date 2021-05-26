Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class ORBStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public StoplossPoint As Decimal
        Public TrailingStartPoint As Decimal
        Public TrailingGap As Decimal
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
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            currentMinuteCandle.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date AndAlso
            Not _parentStrategy.IsTradeOpen(currentMinuteCandle, Trade.TypeOfTrade.MIS) AndAlso
            Not _parentStrategy.IsTradeActive(currentMinuteCandle, Trade.TypeOfTrade.MIS) Then
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandle, Trade.TypeOfTrade.MIS)
            If lastExecutedTrade Is Nothing Then
                Dim signalCandle As Payload = currentMinuteCandle.PreviousCandlePayload
                If currentTick.Open < signalCandle.High Then
                    Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                             .EntryPrice = signalCandle.High,
                             .EntryDirection = Trade.TradeExecutionDirection.Buy,
                             .Quantity = Me.LotSize,
                             .Stoploss = .EntryPrice - _userInputs.StoplossPoint,
                             .Target = .EntryPrice + 100000000000,
                             .Buffer = 0,
                             .SignalCandle = signalCandle,
                             .OrderType = Trade.TypeOfOrder.SL,
                             .Supporting1 = "Stoploss Order"
                         }

                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                Else
                    Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                            .EntryPrice = signalCandle.High,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = Me.LotSize,
                            .Stoploss = .EntryPrice - _userInputs.StoplossPoint,
                            .Target = .EntryPrice + 100000000000,
                            .Buffer = 0,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = "Market Order"
                        }

                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                End If

                If currentTick.Open > signalCandle.Low Then
                    Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                             .EntryPrice = signalCandle.Low,
                             .EntryDirection = Trade.TradeExecutionDirection.Sell,
                             .Quantity = Me.LotSize,
                             .Stoploss = .EntryPrice + _userInputs.StoplossPoint,
                             .Target = .EntryPrice - 100000000000,
                             .Buffer = 0,
                             .SignalCandle = signalCandle,
                             .OrderType = Trade.TypeOfOrder.SL,
                             .Supporting1 = "Stoploss Order"
                         }

                    ret.Item2.Add(parameter)
                Else
                    Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                            .EntryPrice = signalCandle.Low,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = Me.LotSize,
                            .Stoploss = .EntryPrice + _userInputs.StoplossPoint,
                            .Target = .EntryPrice - 100000000000,
                            .Buffer = 0,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = "Market Order"
                        }

                    ret.Item2.Add(parameter)
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandle, Trade.TypeOfTrade.MIS)
            If lastExecutedTrade IsNot Nothing Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Trade")
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim pl As Decimal = currentTick.Open - currentTrade.EntryPrice
                If pl >= _userInputs.TrailingStartPoint Then
                    Dim mul As Integer = Math.Floor(pl / _userInputs.TrailingGap)
                    Dim triggerPrice As Decimal = currentTrade.EntryPrice + _userInputs.TrailingGap * (mul - 1)
                    If currentTrade.PotentialStopLoss < triggerPrice Then
                        ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Trailed at {0}", _userInputs.TrailingGap * (mul - 1)))
                    End If
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim pl As Decimal = currentTrade.EntryPrice - currentTick.Open
                If pl >= _userInputs.TrailingStartPoint Then
                    Dim mul As Integer = Math.Floor(pl / _userInputs.TrailingGap)
                    Dim triggerPrice As Decimal = currentTrade.EntryPrice - _userInputs.TrailingGap * (mul - 1)
                    If currentTrade.PotentialStopLoss > triggerPrice Then
                        ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Trailed at {0}", _userInputs.TrailingGap * (mul - 1)))
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

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function
End Class
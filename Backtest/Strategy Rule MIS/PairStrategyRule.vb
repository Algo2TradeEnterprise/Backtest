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
                   ByVal canceller As CancellationTokenSource,
                   ByVal direction As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, canceller)
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
            Dim parameter As PlaceOrderParameters = Nothing
            If _direction = Trade.TradeExecutionDirection.Buy Then
                parameter = New PlaceOrderParameters With {
                                        .EntryPrice = currentTick.Open,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = quantity,
                                        .Stoploss = .EntryPrice - 100000,
                                        .Target = .EntryPrice + 100000,
                                        .Buffer = 0,
                                        .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting1 = "Force Entry"
                                    }
            ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                parameter = New PlaceOrderParameters With {
                                .EntryPrice = currentTick.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + 100000,
                                .Target = .EntryPrice - 100000,
                                .Buffer = 0,
                                .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = "Force Entry"
                            }
            End If

            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            Me.ForceTakeTrade = False
        Else
            Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
            Dim parameter As PlaceOrderParameters = Nothing
            If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
                Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
                _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
                _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
                _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
                _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
                _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
                _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
                _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
                currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

                Dim signalCandle As Payload = Nothing
                'Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                'If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                'End If
                If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromInvestment(Me.LotSize, 10000, currentTick.Open, _parentStrategy.StockType, True)

                    If _direction = Trade.TradeExecutionDirection.Buy Then
                        parameter = New PlaceOrderParameters With {
                                        .EntryPrice = currentTick.Open,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = quantity,
                                        .Stoploss = .EntryPrice - 100000,
                                        .Target = .EntryPrice + 100000,
                                        .Buffer = 0,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting1 = "Normal Entry"
                                    }
                    ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                        parameter = New PlaceOrderParameters With {
                                        .EntryPrice = currentTick.Open,
                                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                        .Quantity = quantity,
                                        .Stoploss = .EntryPrice + 100000,
                                        .Target = .EntryPrice - 100000,
                                        .Buffer = 0,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting1 = "Normal Entry"
                                    }
                    End If
                End If
            End If
            If parameter IsNot Nothing Then
                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                Me.AnotherPairInstrument.ForceTakeTrade = True
            End If
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
End Class

Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.DAL
Imports System.IO

Public Class NiftyBankniftyPairStrategy
    Inherits StrategyRule


    Public DummyCandle As Payload = Nothing

    Private ReadOnly _direction As Trade.TradeExecutionDirection
    Private ReadOnly _slPoint As Decimal
    Private ReadOnly _numberOfLots As Integer

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal canceller As CancellationTokenSource,
                   ByVal direction As Decimal,
                   ByVal slPoint As Decimal,
                   ByVal numberOfLots As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, canceller)
        If direction > 0 Then
            _direction = Trade.TradeExecutionDirection.Buy
        ElseIf direction < 0 Then
            _direction = Trade.TradeExecutionDirection.Sell
        End If
        _slPoint = slPoint
        _numberOfLots = numberOfLots
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
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

            Me.DummyCandle = currentTick

            Dim signalCandle As Payload = Nothing
            If CType(Me.AnotherPairInstrument, NiftyBankniftyPairStrategy).DummyCandle IsNot Nothing Then
                Dim pairlastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(CType(Me.AnotherPairInstrument, NiftyBankniftyPairStrategy).DummyCandle, Trade.TypeOfTrade.MIS)
                If pairlastExecutedOrder IsNot Nothing Then
                    If pairlastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                        Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(pairlastExecutedOrder.ExitTime, _signalPayload))
                        If currentMinuteCandlePayload.PayloadDate > exitCandle.PayloadDate Then
                            signalCandle = currentMinuteCandlePayload
                        End If
                    ElseIf pairlastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                        Dim entryCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(pairlastExecutedOrder.EntryTime, _signalPayload))
                        If entryCandle.PayloadDate = currentMinuteCandlePayload.PayloadDate Then
                            signalCandle = currentMinuteCandlePayload
                        End If
                    End If
                Else
                    signalCandle = currentMinuteCandlePayload
                End If
            Else
                signalCandle = currentMinuteCandlePayload
            End If
            If signalCandle IsNot Nothing Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
                If lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                    If currentMinuteCandlePayload.PayloadDate < lastExecutedOrder.ExitTime Then
                        signalCandle = Nothing
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim quantity As Decimal = Me.LotSize * _numberOfLots
                If _direction = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalCandle.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - _slPoint,
                                .Target = .EntryPrice + 1000000,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market
                            }
                ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalCandle.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + _slPoint,
                                .Target = .EntryPrice - 1000000,
                                .Buffer = 0,
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
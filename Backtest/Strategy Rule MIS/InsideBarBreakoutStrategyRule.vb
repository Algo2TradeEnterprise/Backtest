Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers

Public Class InsideBarBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private ReadOnly _tradeStartTime As Date
    Private ReadOnly _firstEntryDirection As Trade.TradeExecutionDirection
    Private ReadOnly _buyLevel As Decimal
    Private ReadOnly _sellLevel As Decimal
    Private ReadOnly _buffer As Decimal

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal tradeStartTime As Date,
                   ByVal direction As Integer,
                   ByVal buyLevel As Decimal,
                   ByVal sellLevel As Decimal,
                   ByVal buffer As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities

        _tradeStartTime = tradeStartTime
        If direction > 0 Then
            _firstEntryDirection = Trade.TradeExecutionDirection.Buy
        ElseIf direction < 0 Then
            _firstEntryDirection = Trade.TradeExecutionDirection.Sell
        End If
        _buyLevel = buyLevel
        _sellLevel = sellLevel
        _buffer = buffer
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        'Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        Dim tradeStartTime As Date = _tradeStartTime

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, _parentStrategy.TradeType) AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Trade.TradeExecutionDirection, Integer, Decimal) = GetSignalForEntry(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
            End If

            If signalCandle IsNot Nothing Then
                If signal.Item2 = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = _buyLevel,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = signal.Item3,
                                .Stoploss = _sellLevel,
                                .Target = .EntryPrice + signal.Item4,
                                .Buffer = _buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL
                            }
                ElseIf signal.Item2 = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = _sellLevel,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = signal.Item3,
                                .Stoploss = _buyLevel,
                                .Target = .EntryPrice - signal.Item4,
                                .Buffer = _buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL
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
            If _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, _parentStrategy.TradeType) Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Trade")
            End If
        End If
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

    Private Function GetSignalForEntry(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection, Integer, Decimal)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection, Integer, Decimal) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            If _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) = 0 Then
                If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) Then
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, _buyLevel, _sellLevel, _userInputs.MaxLossPerTrade, _parentStrategy.StockType)
                    Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, _buyLevel, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                    Dim targetPoint As Decimal = targetPrice - _buyLevel
                    ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Integer, Decimal)(True, _firstEntryDirection, quantity, targetPoint)
                End If
            ElseIf _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) = 1 Then
                Dim directionToTrade As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
                Dim lastTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
                If lastTrade IsNot Nothing Then
                    If lastTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        directionToTrade = Trade.TradeExecutionDirection.Sell
                    ElseIf lastTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        directionToTrade = Trade.TradeExecutionDirection.Buy
                    End If
                    If directionToTrade <> Trade.TradeExecutionDirection.None Then
                        If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, directionToTrade) AndAlso
                            Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, directionToTrade) Then
                            Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, _buyLevel, lastTrade.Quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                            Dim targetPoint As Decimal = targetPrice - _buyLevel
                            ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Integer, Decimal)(True, directionToTrade, lastTrade.Quantity * 2, targetPoint)
                        End If
                    End If
                End If
            ElseIf _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) = 2 Then
                If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) Then
                    Dim lastTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
                    If lastTrade IsNot Nothing Then
                        Dim quantity As Integer = lastTrade.Quantity / 2
                        Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, _buyLevel, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                        Dim targetPoint As Decimal = targetPrice - _buyLevel
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Integer, Decimal)(True, _firstEntryDirection, quantity * 3, targetPoint)
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class
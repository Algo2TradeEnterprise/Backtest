Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class BollingerCloseStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public ATRMultiplier As Decimal
        Public BollingerPeriod As Decimal
        Public StandardDeviation As Decimal
        Public ExitAtProfit As Boolean
    End Class
#End Region

    Private _bollingerHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _smaPayload As Dictionary(Of Date, Decimal) = Nothing
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
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.BollingerBands.CalculateBollingerBands(_userInputs.BollingerPeriod, Payload.PayloadFields.Close, _userInputs.StandardDeviation, _signalPayload, _bollingerHighPayload, _bollingerLowPayload, _smaPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso currentMinuteCandlePayload.PayloadDate >= _tradeStartTime Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item4
            End If

            If signalCandle IsNot Nothing Then
                Dim atr As Decimal = _atrPayload(signalCandle.PayloadDate)
                If signal.Item5 = Trade.TradeExecutionDirection.Buy AndAlso (Not _userInputs.ExitAtProfit OrElse
                    (_parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) OrElse
                     (Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                      _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol) < Math.Abs(_userInputs.MaxLossPerTrade)))) Then

                    Dim entryPrice As Decimal = currentTick.Open
                    Dim quantity As Integer = signal.Item2

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = entryPrice - 100000000,
                                    .Target = entryPrice + 100000000,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item3,
                                    .Supporting3 = atr
                                }
                ElseIf signal.Item5 = Trade.TradeExecutionDirection.Sell AndAlso (Not _userInputs.ExitAtProfit OrElse
                    (_parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) OrElse
                     (Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                      _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol) < Math.Abs(_userInputs.MaxLossPerTrade)))) Then

                    Dim entryPrice As Decimal = currentTick.Open
                    Dim quantity As Integer = signal.Item2

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = entryPrice + 100000000,
                                    .Target = entryPrice - 100000000,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item3,
                                    .Supporting3 = atr
                                }
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter IsNot Nothing Then
            parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim bollingerHigh As Decimal = _bollingerHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
            Dim bollingerLow As Decimal = _bollingerLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso currentMinuteCandlePayload.PreviousCandlePayload.Close > bollingerHigh Then
                ret = New Tuple(Of Boolean, String)(True, "Reverse Signal")
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso currentMinuteCandlePayload.PreviousCandlePayload.Close < bollingerLow Then
                ret = New Tuple(Of Boolean, String)(True, "Reverse Signal")
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim bollingerHigh As Decimal = _bollingerHighPayload(candle.PayloadDate)
            Dim bollingerLow As Decimal = _bollingerLowPayload(candle.PayloadDate)
            Dim atr As Decimal = _atrPayload(candle.PayloadDate) * _userInputs.ATRMultiplier
            Dim quantity As Integer = Me.LotSize
            If _parentStrategy.StockType = Trade.TypeOfStock.Cash Then
                quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, currentTick.Open, currentTick.Open - atr, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
            End If
            Dim iteration As Integer = 1
            If candle.Close < bollingerLow Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(candle, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy)
                If lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                    If lastExecutedOrder.EntryPrice - candle.Close >= atr Then
                        iteration = Val(lastExecutedOrder.Supporting2) + 1
                        quantity = lastExecutedOrder.Quantity + (lastExecutedOrder.Quantity / (iteration - 1))
                        ret = New Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection)(True, quantity, iteration, candle, Trade.TradeExecutionDirection.Buy)
                    End If
                Else
                    ret = New Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection)(True, quantity, iteration, candle, Trade.TradeExecutionDirection.Buy)
                End If
            ElseIf candle.Close > bollingerHigh Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(candle, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Sell)
                If lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                    If candle.Close - lastExecutedOrder.EntryPrice >= atr Then
                        iteration = Val(lastExecutedOrder.Supporting2) + 1
                        quantity = lastExecutedOrder.Quantity + (lastExecutedOrder.Quantity / (iteration - 1))
                        ret = New Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection)(True, quantity, iteration, candle, Trade.TradeExecutionDirection.Sell)
                    End If
                Else
                    ret = New Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection)(True, quantity, iteration, candle, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function
End Class
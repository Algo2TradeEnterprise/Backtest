Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class SmallRangeBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
        Public BreakevenMovement As Boolean
        Public ReverseSignalEntry As Boolean
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _signalCandle As Payload = Nothing

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
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS) Then

            If ((_userInputs.ReverseSignalEntry AndAlso Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy)) OrElse
                (Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS))) Then
                Dim signalCandle As Payload = Nothing
                Dim signal As Tuple(Of Boolean, Payload) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy)
                    If lastExecutedOrder Is Nothing Then
                        signalCandle = signal.Item2
                    ElseIf lastExecutedOrder IsNot Nothing Then
                        If lastExecutedOrder.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                            signalCandle = signal.Item2
                        End If
                    End If
                End If
                If signalCandle IsNot Nothing Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandle.High + buffer
                    Dim stoploss As Decimal = signalCandle.Low - buffer
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)
                    If currentTick.Open < entryPrice Then
                        parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = Math.Round((signalCandle.CandleRange / _atrPayload(signalCandle.PayloadDate)) * 100, 4)
                                }
                    End If
                End If
            End If
            If ((_userInputs.ReverseSignalEntry AndAlso Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell)) OrElse
                (Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS))) Then
                Dim signalCandle As Payload = Nothing
                Dim signal As Tuple(Of Boolean, Payload) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Sell)
                    If lastExecutedOrder Is Nothing Then
                        signalCandle = signal.Item2
                    ElseIf lastExecutedOrder IsNot Nothing Then
                        If lastExecutedOrder.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                            signalCandle = signal.Item2
                        End If
                    End If
                End If
                If signalCandle IsNot Nothing Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandle.Low - buffer
                    Dim stoploss As Decimal = signalCandle.High + buffer
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Cash)
                    If currentTick.Open > entryPrice Then
                        parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = Math.Round((signalCandle.CandleRange / _atrPayload(signalCandle.PayloadDate)) * 100, 4)
                                }
                    End If
                End If
            End If

        End If
        Dim parameterList As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            If parameterList Is Nothing Then parameterList = New List(Of PlaceOrderParameters)
            parameterList.Add(parameter1)
        End If
        If parameter2 IsNot Nothing Then
            If parameterList Is Nothing Then parameterList = New List(Of PlaceOrderParameters)
            parameterList.Add(parameter2)
        End If
        If parameterList IsNot Nothing AndAlso parameterList.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameterList)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            If _parentStrategy.StockNumberOfTrades(_tradingDate.Date, _tradingSymbol) >= _parentStrategy.NumberOfTradesPerStockPerDay Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            End If
            Dim signal As Tuple(Of Boolean, Payload) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, currentTrade.EntryDirection)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
            If _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS) Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            End If
        ElseIf currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If _parentStrategy.StockNumberOfTrades(_tradingDate.Date, _tradingSymbol) = _parentStrategy.NumberOfTradesPerStockPerDay Then
                If _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol) >= 0 Then
                    ret = New Tuple(Of Boolean, String)(True, "Loss makeup done")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _userInputs.BreakevenMovement AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim slPoint As Decimal = currentTrade.SignalCandle.CandleRange + 2 * currentTrade.StoplossBuffer
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentTick.Open >= currentTrade.EntryPrice + slPoint Then
                    triggerPrice = currentTrade.EntryPrice + _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Trade.TradeExecutionDirection.Buy, Me.LotSize, _parentStrategy.StockType)
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentTick.Open <= currentTrade.EntryPrice - slPoint Then
                    triggerPrice = currentTrade.EntryPrice - _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Trade.TradeExecutionDirection.Sell, Me.LotSize, _parentStrategy.StockType)
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("({0})Breakeven at {1}", slPoint, currentTick.PayloadDate.ToString("HH:mm:ss")))
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Payload)
        Dim ret As Tuple(Of Boolean, Payload) = Nothing
        If candle IsNot Nothing Then
            Dim atr As Decimal = _atrPayload(candle.PayloadDate)
            If candle.CandleRange < atr AndAlso candle.CandleRange > 0.5 * atr Then
                _signalCandle = candle
            End If
            '_signalCandle = _signalPayload(New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, 9, 15, 0))
        End If
        If _signalCandle IsNot Nothing Then
            ret = New Tuple(Of Boolean, Payload)(True, _signalCandle)
        End If
        Return ret
    End Function
End Class
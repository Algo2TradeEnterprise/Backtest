Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HighestLowestPointStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public ATRMultiplier As Decimal
        Public MaxLossPerTrade As Decimal
    End Class
#End Region
    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _previousDayLastCandleATR As Decimal

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
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            For Each runningPayload In _signalPayload
                If runningPayload.Value.PayloadDate.Date = _tradingDate.Date AndAlso
                    runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                    _previousDayLastCandleATR = _atrPayload(runningPayload.Value.PreviousCandlePayload.PayloadDate)
                    Exit For
                End If
            Next
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signal As Tuple(Of Boolean, Payload) = GetEntrySignal(currentMinuteCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy)
                    If lastExecutedOrder Is Nothing OrElse lastExecutedOrder.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2.Low, RoundOfType.Floor)
                        Dim atrPoint As Decimal = ConvertFloorCeling(_previousDayLastCandleATR * _userInputs.ATRMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                        Dim entryPrice As Decimal = signal.Item2.Low + atrPoint
                        Dim stoploss As Decimal = signal.Item2.Low - buffer
                        Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, _userInputs.MaxLossPerTrade, _parentStrategy.StockType)

                        If currentTick.Open < entryPrice Then
                            parameter1 = New PlaceOrderParameters With {
                                        .EntryPrice = entryPrice,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = quantity,
                                        .Stoploss = stoploss,
                                        .Target = entryPrice + 100000000000,
                                        .Buffer = buffer,
                                        .SignalCandle = signal.Item2,
                                        .OrderType = Trade.TypeOfOrder.SL,
                                        .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                        .Supporting2 = atrPoint
                                    }
                        End If
                    End If
                End If
            End If

            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signal As Tuple(Of Boolean, Payload) = GetEntrySignal(currentMinuteCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Sell)
                    If lastExecutedOrder Is Nothing OrElse lastExecutedOrder.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2.High, RoundOfType.Floor)
                        Dim atrPoint As Decimal = ConvertFloorCeling(_previousDayLastCandleATR * _userInputs.ATRMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                        Dim entryPrice As Decimal = signal.Item2.High - atrPoint
                        Dim stoploss As Decimal = signal.Item2.High + buffer
                        Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, _userInputs.MaxLossPerTrade, _parentStrategy.StockType)

                        If currentTick.Open > entryPrice Then
                            parameter2 = New PlaceOrderParameters With {
                                        .EntryPrice = entryPrice,
                                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                        .Quantity = quantity,
                                        .Stoploss = stoploss,
                                        .Target = entryPrice - 100000000000,
                                        .Buffer = buffer,
                                        .SignalCandle = signal.Item2,
                                        .OrderType = Trade.TypeOfOrder.SL,
                                        .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                        .Supporting2 = atrPoint
                                    }
                        End If
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Payload) = GetEntrySignal(currentMinuteCandlePayload, currentTick, currentTrade.EntryDirection)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Payload)
        Dim ret As Tuple(Of Boolean, Payload) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            If direction = Trade.TradeExecutionDirection.Sell Then
                Dim highestCandle As Payload = GetHighestCandleOfTheDay(currentCandle, currentTick)
                If highestCandle IsNot Nothing Then
                    ret = New Tuple(Of Boolean, Payload)(True, highestCandle)
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Buy Then
                Dim lowestCandle As Payload = GetLowestCandleOfTheDay(currentCandle, currentTick)
                If lowestCandle IsNot Nothing Then
                    ret = New Tuple(Of Boolean, Payload)(True, lowestCandle)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestCandleOfTheDay(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Payload
        Dim ret As Payload = Nothing
        For Each runningCandle In _signalPayload
            If runningCandle.Key.Date = _tradingDate.Date AndAlso runningCandle.Key < currentCandle.PayloadDate Then
                If ret Is Nothing Then
                    ret = runningCandle.Value
                ElseIf ret IsNot Nothing AndAlso ret.High <= runningCandle.Value.High Then
                    ret = runningCandle.Value
                End If
            End If
        Next
        If currentCandle.Ticks IsNot Nothing AndAlso currentCandle.Ticks.Count > 0 Then
            For Each runningTick In currentCandle.Ticks
                If runningTick.PayloadDate < currentTick.PayloadDate Then
                    If ret Is Nothing Then
                        ret = runningTick
                    ElseIf ret IsNot Nothing AndAlso ret.High <= runningTick.High Then
                        ret = runningTick
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Private Function GetLowestCandleOfTheDay(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Payload
        Dim ret As Payload = Nothing
        For Each runningCandle In _signalPayload
            If runningCandle.Key.Date = _tradingDate.Date AndAlso runningCandle.Key < currentCandle.PayloadDate Then
                If ret Is Nothing Then
                    ret = runningCandle.Value
                ElseIf ret IsNot Nothing AndAlso ret.Low >= runningCandle.Value.Low Then
                    ret = runningCandle.Value
                End If
            End If
        Next
        If currentCandle.Ticks IsNot Nothing AndAlso currentCandle.Ticks.Count > 0 Then
            For Each runningTick In currentCandle.Ticks
                If runningTick.PayloadDate < currentTick.PayloadDate Then
                    If ret Is Nothing Then
                        ret = runningTick
                    ElseIf ret IsNot Nothing AndAlso ret.Low >= runningTick.Low Then
                        ret = runningTick
                    End If
                End If
            Next
        End If
        Return ret
    End Function
End Class
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class EmaSmaCrossoverStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _dayATR As Decimal
    Private ReadOnly _slPoint As Decimal
    Private ReadOnly _slab As Decimal

    Private _ema13Payload As Dictionary(Of Date, Decimal) = Nothing
    Private _sma50Payload As Dictionary(Of Date, Decimal) = Nothing
    Private _signalCheckStartTime As Date

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal dayATR As Decimal,
                   ByVal slab As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        _dayATR = dayATR
        _slab = slab
        '_slPoint = ConvertFloorCeling(_dayATR / 8, _parentStrategy.TickSize, RoundOfType.Floor)
        _slPoint = _slab
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.EMA.CalculateEMA(13, Payload.PayloadFields.Close, _signalPayload, _ema13Payload)
        Indicator.SMA.CalculateSMA(50, Payload.PayloadFields.Close, _signalPayload, _sma50Payload)

        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = _tradingDate.Date Then
                Dim ema13 As Decimal = _ema13Payload(runningPayload.Key)
                Dim sma50 As Decimal = _sma50Payload(runningPayload.Key)
                If (ema13 >= runningPayload.Value.Low AndAlso ema13 <= runningPayload.Value.High) OrElse
                    (sma50 >= runningPayload.Value.Low AndAlso sma50 <= runningPayload.Value.High) Then
                    _signalCheckStartTime = runningPayload.Key
                    Exit For
                End If
            End If
        Next
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
                Dim signalCandle As Payload = Nothing
                Dim signal As Tuple(Of Boolean, Payload, String) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    signalCandle = signal.Item2
                End If
                If signalCandle IsNot Nothing Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandle.High + buffer
                    Dim stoploss As Decimal = entryPrice - _slPoint
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = entryPrice + 100000000000
                    If _userInputs.TargetMultiplier <> Decimal.MaxValue Then
                        target = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)
                    End If
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
                                    .Supporting2 = _slPoint,
                                    .Supporting3 = signal.Item3
                                }
                    End If
                End If
            End If
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signalCandle As Payload = Nothing
                Dim signal As Tuple(Of Boolean, Payload, String) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    signalCandle = signal.Item2
                End If
                If signalCandle IsNot Nothing Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandle.Low - buffer
                    Dim stoploss As Decimal = entryPrice + _slPoint
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = entryPrice - 100000000000
                    If _userInputs.TargetMultiplier <> Decimal.MaxValue Then
                        target = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Cash)
                    End If
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
                                    .Supporting2 = _slPoint,
                                    .Supporting3 = signal.Item3
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
            Dim ema13 As Decimal = _ema13Payload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
            Dim sma50 As Decimal = _sma50Payload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso ema13 < sma50 Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso ema13 > sma50 Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            End If
            If ret Is Nothing Then
                Dim signal As Tuple(Of Boolean, Payload, String) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, currentTrade.EntryDirection)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    If currentTrade.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
            End If
        ElseIf currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso currentTrade.Supporting3 = "Normal" Then
            Dim ema13 As Decimal = _ema13Payload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
            Dim sma50 As Decimal = _sma50Payload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso ema13 < sma50 Then
                ret = New Tuple(Of Boolean, String)(True, "EMA Crossover")
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso ema13 > sma50 Then
                ret = New Tuple(Of Boolean, String)(True, "EMA Crossover")
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Payload, String)
        Dim ret As Tuple(Of Boolean, Payload, String) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing AndAlso candle.PayloadDate >= _signalCheckStartTime Then
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(candle, Trade.TypeOfTrade.MIS, direction)
            Dim startTime As Date = _tradingDate.Date
            If lastExecutedTrade IsNot Nothing Then startTime = lastExecutedTrade.ExitTime
            If IsEligibleToTakeTrade(startTime, candle.PayloadDate, direction) Then
                Dim ema13 As Decimal = _ema13Payload(candle.PayloadDate)
                Dim sma50 As Decimal = _sma50Payload(candle.PayloadDate)
                If direction = Trade.TradeExecutionDirection.Buy AndAlso ema13 > sma50 Then
                    If candle.High < candle.PreviousCandlePayload.High Then
                        Dim remark As String = "Normal"
                        If candle.High < sma50 Then remark = "Special"
                        ret = New Tuple(Of Boolean, Payload, String)(True, candle, remark)
                    End If
                ElseIf direction = Trade.TradeExecutionDirection.Sell AndAlso ema13 < sma50 Then
                    If candle.Low > candle.PreviousCandlePayload.Low Then
                        Dim remark As String = "Normal"
                        If candle.Low > sma50 Then remark = "Special"
                        ret = New Tuple(Of Boolean, Payload, String)(True, candle, remark)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsEligibleToTakeTrade(ByVal startTime As Date, ByVal endTime As Date, ByVal direction As Trade.TradeExecutionDirection) As Boolean
        Dim ret As Boolean = False
        For Each runningPayload In _signalPayload.OrderByDescending(Function(x)
                                                                        Return x.Key
                                                                    End Function)
            If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key > startTime AndAlso runningPayload.Key <= endTime Then
                Dim ema As Decimal = _ema13Payload(runningPayload.Key)
                Dim sma As Decimal = _sma50Payload(runningPayload.Key)
                If direction = Trade.TradeExecutionDirection.Buy AndAlso runningPayload.Value.Close < ema AndAlso ema > sma Then
                    ret = True
                    Exit For
                ElseIf direction = Trade.TradeExecutionDirection.Sell AndAlso runningPayload.Value.Close > ema AndAlso ema < sma Then
                    ret = True
                    Exit For
                End If
                If direction = Trade.TradeExecutionDirection.Buy AndAlso ema < sma Then
                    Exit For
                ElseIf direction = Trade.TradeExecutionDirection.Sell AndAlso ema > sma Then
                    Exit For
                End If
            End If
        Next
        Return ret
    End Function
End Class
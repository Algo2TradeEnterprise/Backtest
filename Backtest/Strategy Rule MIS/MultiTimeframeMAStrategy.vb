Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class MultiTimeframeMAStrategy
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public HigherTimeframe As Integer
        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
        Public TargetInINR As Boolean
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _xMinPayload As Dictionary(Of Date, Payload)
    Private _htSMAPayload As Dictionary(Of Date, Decimal)
    Private _ltSMAPayload As Dictionary(Of Date, Decimal)

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

        Dim exchangeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day,
                                                             _parentStrategy.ExchangeStartTime.Hours, _parentStrategy.ExchangeStartTime.Minutes, _parentStrategy.ExchangeStartTime.Seconds)
        _xMinPayload = Common.ConvertPayloadsToXMinutes(_inputPayload, _userInputs.HigherTimeframe, exchangeStartTime)
        Indicator.SMA.CalculateSMA(20, Payload.PayloadFields.Close, _xMinPayload, _htSMAPayload)
        Indicator.SMA.CalculateSMA(20, Payload.PayloadFields.Close, _signalPayload, _ltSMAPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandle, _parentStrategy.TradeType)
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item3
                ElseIf lastExecutedOrder IsNot Nothing Then
                    Dim exitPayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedOrder.ExitTime, _signalPayload))
                    If signal.Item3.PayloadDate >= exitPayload.PayloadDate Then
                        signalCandle = signal.Item3
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = signal.Item2 + buffer
                    Dim stoploss As Decimal = signalCandle.Low - buffer
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade * _userInputs.TargetMultiplier), Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)
                    If Not _userInputs.TargetInINR Then
                        target = entryPrice + (entryPrice - stoploss) * _userInputs.TargetMultiplier
                    End If

                    parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = signal.Item2 - buffer
                    Dim stoploss As Decimal = signalCandle.High + buffer
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade * _userInputs.TargetMultiplier), Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Cash)
                    If Not _userInputs.TargetInINR Then
                        target = entryPrice - (stoploss - entryPrice) * _userInputs.TargetMultiplier
                    End If

                    parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter1)
            If parameter2 IsNot Nothing Then parameters.Add(parameter2)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing Then
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                    currentMinuteCandle.PreviousCandlePayload.High < _ltSMAPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                    currentMinuteCandle.PreviousCandlePayload.Low > _ltSMAPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim direction As Trade.TradeExecutionDirection = GetDirection(candle)
            If direction = Trade.TradeExecutionDirection.Buy Then
                If candle.High >= Math.Round(_ltSMAPayload(candle.PayloadDate), 2) AndAlso candle.Low <= Math.Round(_ltSMAPayload(candle.PayloadDate), 2) AndAlso
                    (candle.PreviousCandlePayload.High < Math.Round(_ltSMAPayload(candle.PreviousCandlePayload.PayloadDate), 2) OrElse
                    candle.PreviousCandlePayload.Low > Math.Round(_ltSMAPayload(candle.PreviousCandlePayload.PayloadDate), 2)) Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, candle.High, candle, Trade.TradeExecutionDirection.Buy)
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                If candle.High >= Math.Round(_ltSMAPayload(candle.PayloadDate), 2) AndAlso candle.Low <= Math.Round(_ltSMAPayload(candle.PayloadDate), 2) AndAlso
                    (candle.PreviousCandlePayload.High < Math.Round(_ltSMAPayload(candle.PreviousCandlePayload.PayloadDate), 2) OrElse
                    candle.PreviousCandlePayload.Low > Math.Round(_ltSMAPayload(candle.PreviousCandlePayload.PayloadDate), 2)) Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, candle.Low, candle, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetDirection(ByVal candle As Payload) As Trade.TradeExecutionDirection
        Dim ret As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
        Dim currentXMinuteCandle As Payload = _xMinPayload(_parentStrategy.GetCurrentXMinuteCandleTime(candle.PayloadDate, _xMinPayload, _userInputs.HigherTimeframe))
        If currentXMinuteCandle IsNot Nothing AndAlso currentXMinuteCandle.PreviousCandlePayload IsNot Nothing Then
            For Each runningPayload In _xMinPayload
                If runningPayload.Key <= currentXMinuteCandle.PreviousCandlePayload.PayloadDate Then
                    If runningPayload.Value.Low >= Math.Round(_htSMAPayload(runningPayload.Key), 2) Then
                        ret = Trade.TradeExecutionDirection.Buy
                    ElseIf runningPayload.Value.High <= Math.Round(_htSMAPayload(runningPayload.Key), 2) Then
                        ret = Trade.TradeExecutionDirection.Sell
                    End If
                End If
            Next
        End If
        Return ret
    End Function
End Class
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class EMAScalpingStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public MaxLossPerTrade As Decimal
        Public BreakevenMovement As Boolean
    End Class
#End Region

    Private _ema50Payload As Dictionary(Of Date, Decimal) = Nothing
    Private _ema100Payload As Dictionary(Of Date, Decimal) = Nothing
    Private _ema150Payload As Dictionary(Of Date, Decimal) = Nothing
    'Private _swingHighPayload As Dictionary(Of Date, Decimal) = Nothing
    'Private _swingLowPayload As Dictionary(Of Date, Decimal) = Nothing

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

        Indicator.EMA.CalculateEMA(50, Payload.PayloadFields.Close, _signalPayload, _ema50Payload)
        Indicator.EMA.CalculateEMA(100, Payload.PayloadFields.Close, _signalPayload, _ema100Payload)
        Indicator.EMA.CalculateEMA(150, Payload.PayloadFields.Close, _signalPayload, _ema150Payload)
        'Indicator.SwingHighLow.CalculateSwingHighLow(_signalPayload, False, _swingHighPayload, _swingLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing

        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item4
                ElseIf lastExecutedOrder IsNot Nothing Then
                    Dim lastOrderExitTime As Date = _parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedOrder.ExitTime, _signalPayload)
                    If signal.Item2.PayloadDate >= lastOrderExitTime Then
                        signalCandle = signal.Item4
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim stoploss As Decimal = Math.Min(signalCandle.Low, signalCandle.PreviousCandlePayload.Low) - buffer
                    If stoploss < entryPrice Then
                        Dim target As Decimal = entryPrice + ConvertFloorCeling((entryPrice - stoploss) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                        Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)

                        If currentTick.Open > _ema50Payload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) Then
                            parameter = New PlaceOrderParameters With {
                                        .EntryPrice = entryPrice,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = quantity,
                                        .Stoploss = stoploss,
                                        .Target = target,
                                        .Buffer = buffer,
                                        .SignalCandle = signal.Item2,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                        .Supporting2 = signal.Item2.PayloadDate.ToString("HH:mm:ss")
                                    }
                        End If
                    End If
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim stoploss As Decimal = Math.Max(signalCandle.High, signalCandle.PreviousCandlePayload.High) + buffer
                    If stoploss > entryPrice Then
                        Dim target As Decimal = entryPrice - ConvertFloorCeling((stoploss - entryPrice) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                        Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)

                        If currentTick.Open < _ema50Payload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) Then
                            parameter = New PlaceOrderParameters With {
                                        .EntryPrice = entryPrice,
                                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                        .Quantity = quantity,
                                        .Stoploss = stoploss,
                                        .Target = target,
                                        .Buffer = buffer,
                                        .SignalCandle = signal.Item2,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                        .Supporting2 = signal.Item2.PayloadDate.ToString("HH:mm:ss")
                                    }
                        End If
                    End If
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
        If _userInputs.BreakevenMovement AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim slPoint As Decimal = Math.Round(Math.Abs(currentTrade.EntryPrice - currentTrade.PotentialStopLoss), 2)
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Payload)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Payload) = Nothing
        If candle IsNot Nothing Then
            If _ema50Payload(candle.PayloadDate) > _ema100Payload(candle.PayloadDate) AndAlso
                _ema100Payload(candle.PayloadDate) > _ema150Payload(candle.PayloadDate) Then
                If candle.Close > _ema50Payload(candle.PayloadDate) Then
                    Dim previousSignal As Tuple(Of Boolean, Payload) = IsPreviousConditionSatisfied(candle.PayloadDate, Trade.TradeExecutionDirection.Buy)
                    If previousSignal IsNot Nothing AndAlso previousSignal.Item1 Then
                        ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Payload)(True, previousSignal.Item2, Trade.TradeExecutionDirection.Buy, candle)
                    End If
                End If
            ElseIf _ema50Payload(candle.PayloadDate) < _ema100Payload(candle.PayloadDate) AndAlso
                _ema100Payload(candle.PayloadDate) < _ema150Payload(candle.PayloadDate) Then
                If candle.Close < _ema50Payload(candle.PayloadDate) Then
                    Dim previousSignal As Tuple(Of Boolean, Payload) = IsPreviousConditionSatisfied(candle.PayloadDate, Trade.TradeExecutionDirection.Sell)
                    If previousSignal IsNot Nothing AndAlso previousSignal.Item1 Then
                        ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Payload)(True, previousSignal.Item2, Trade.TradeExecutionDirection.Sell, candle)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsPreviousConditionSatisfied(ByVal beforeThisTime As Date, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Payload)
        Dim ret As Tuple(Of Boolean, Payload) = Nothing
        Dim middleSignal As Payload = Nothing
        For Each runningPayload In _signalPayload.OrderByDescending(Function(x)
                                                                        Return x.Key
                                                                    End Function)
            _cts.Token.ThrowIfCancellationRequested()
            If runningPayload.Key < beforeThisTime AndAlso runningPayload.Key.Date = _tradingDate.Date Then
                If direction = Trade.TradeExecutionDirection.Buy Then
                    If _ema50Payload(runningPayload.Key) > _ema100Payload(runningPayload.Key) AndAlso
                        _ema100Payload(runningPayload.Key) > _ema150Payload(runningPayload.Key) Then
                        If runningPayload.Value.Close > _ema150Payload(runningPayload.Value.PayloadDate) Then
                            If runningPayload.Value.Close < _ema50Payload(runningPayload.Value.PayloadDate) Then
                                middleSignal = runningPayload.Value
                            End If
                            If middleSignal IsNot Nothing Then
                                If runningPayload.Value.Close > _ema50Payload(runningPayload.Value.PayloadDate) Then
                                    ret = New Tuple(Of Boolean, Payload)(True, runningPayload.Value)
                                    Exit For
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                    If _ema50Payload(runningPayload.Key) < _ema100Payload(runningPayload.Key) AndAlso
                        _ema100Payload(runningPayload.Key) < _ema150Payload(runningPayload.Key) Then
                        If runningPayload.Value.Close < _ema150Payload(runningPayload.Value.PayloadDate) Then
                            If runningPayload.Value.Close > _ema50Payload(runningPayload.Value.PayloadDate) Then
                                middleSignal = runningPayload.Value
                            End If
                            If middleSignal IsNot Nothing Then
                                If runningPayload.Value.Close < _ema50Payload(runningPayload.Value.PayloadDate) Then
                                    ret = New Tuple(Of Boolean, Payload)(True, runningPayload.Value)
                                    Exit For
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                End If
            End If
        Next
        Return ret
    End Function
End Class
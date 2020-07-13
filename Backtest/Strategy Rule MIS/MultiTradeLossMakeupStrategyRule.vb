Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class MultiTradeLossMakeupStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
        Public BreakevenMovement As Boolean
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _emaPayload As Dictionary(Of Date, Decimal) = Nothing

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
        Indicator.EMA.CalculateEMA(13, Payload.PayloadFields.Close, _signalPayload, _emaPayload)
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
                Dim signal As Tuple(Of Boolean, Payload) = GetSignal(currentMinuteCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2.High, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2.High + buffer
                    Dim stoploss As Decimal = signal.Item2.Low - buffer
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
                                    .SignalCandle = signal.Item2,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = entryPrice - stoploss
                                }
                    End If
                End If
            End If
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signal As Tuple(Of Boolean, Payload) = GetSignal(currentMinuteCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2.Low, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2.Low - buffer
                    Dim stoploss As Decimal = signal.Item2.High + buffer
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
                                    .SignalCandle = signal.Item2,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = stoploss - entryPrice
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
            Dim signal As Tuple(Of Boolean, Payload) = GetSignal(currentMinuteCandlePayload, currentTick, currentTrade.EntryDirection)
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
        If _userInputs.BreakevenMovement AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim slPoint As Decimal = currentTrade.Supporting2
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

    Private Function GetSignal(ByVal currentCandle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Payload)
        Dim ret As Tuple(Of Boolean, Payload) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing AndAlso Not currentCandle.PreviousCandlePayload.DeadCandle Then
            If Not _parentStrategy.IsTradeActive(currentCandle, Trade.TypeOfTrade.MIS) Then
                Dim signalCandle As Payload = Nothing
                Dim atr As Decimal = _atrPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                Dim ema As Decimal = _emaPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                If currentCandle.PreviousCandlePayload.Low < ema AndAlso currentCandle.PreviousCandlePayload.High > ema Then
                    If currentCandle.PreviousCandlePayload.CandleRange < atr AndAlso currentCandle.PreviousCandlePayload.CandleRange >= atr / 2 Then
                        signalCandle = currentCandle.PreviousCandlePayload
                    End If
                End If

                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS)
                If signalCandle Is Nothing Then
                    If lastExecutedOrder IsNot Nothing Then signalCandle = lastExecutedOrder.SignalCandle
                End If

                If signalCandle IsNot Nothing Then
                    If direction = Trade.TradeExecutionDirection.Buy Then
                        If Not (lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                            lastExecutedOrder.SignalCandle.PayloadDate = signalCandle.PayloadDate) Then
                            ret = New Tuple(Of Boolean, Payload)(True, signalCandle)
                        End If
                    ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                        If Not (lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                            lastExecutedOrder.SignalCandle.PayloadDate = signalCandle.PayloadDate) Then
                            ret = New Tuple(Of Boolean, Payload)(True, signalCandle)
                        End If
                    End If
                End If
            Else
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS)
                Dim signalCandle As Payload = lastExecutedOrder.SignalCandle

                If signalCandle IsNot Nothing Then
                    If direction = Trade.TradeExecutionDirection.Buy Then
                        If Not (lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                            lastExecutedOrder.SignalCandle.PayloadDate = signalCandle.PayloadDate) Then
                            ret = New Tuple(Of Boolean, Payload)(True, signalCandle)
                        End If
                    ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                        If Not (lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                            lastExecutedOrder.SignalCandle.PayloadDate = signalCandle.PayloadDate) Then
                            ret = New Tuple(Of Boolean, Payload)(True, signalCandle)
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class
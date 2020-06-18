Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class MomentumReversalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public MaxLossPerTrade As Decimal
        Public BreakevenMovement As Boolean
    End Class
#End Region

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
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item3
                ElseIf lastExecutedOrder IsNot Nothing Then
                    If signal.Item3.PayloadDate <> lastExecutedOrder.SignalCandle.PayloadDate Then
                        signalCandle = signal.Item3
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim atr As Decimal = ConvertFloorCeling(_atrPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = signal.Item2
                    Dim stoploss As Decimal = signalCandle.Low - buffer
                    Dim slRemark As String = "CR"
                    If signalCandle.CandleRange < atr / 2 Then
                        stoploss = entryPrice - ConvertFloorCeling(atr / 2, _parentStrategy.TickSize, RoundOfType.Celing)
                        slRemark = "Half ATR"
                    Else
                        If entryPrice - atr > stoploss Then
                            stoploss = entryPrice - atr
                            slRemark = "ATR"
                        End If
                    End If
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade * _userInputs.TargetMultiplier), Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = slRemark,
                                    .Supporting3 = (entryPrice - stoploss)
                                }
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = signal.Item2
                    Dim stoploss As Decimal = signalCandle.High + buffer
                    Dim slRemark As String = "CR"
                    If signalCandle.CandleRange < atr / 2 Then
                        stoploss = entryPrice + ConvertFloorCeling(atr / 2, _parentStrategy.TickSize, RoundOfType.Celing)
                        slRemark = "Half ATR"
                    Else
                        If entryPrice + atr < stoploss Then
                            stoploss = entryPrice + atr
                            slRemark = "ATR"
                        End If
                    End If
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade * _userInputs.TargetMultiplier), Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Cash)

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = slRemark,
                                    .Supporting3 = (stoploss - entryPrice)
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
            If ret Is Nothing Then
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                    currentMinuteCandlePayload.PreviousCandlePayload.Close < currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                    currentMinuteCandlePayload.PreviousCandlePayload.Close > currentTrade.PotentialStopLoss Then
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
            Dim slPoint As Decimal = currentTrade.Supporting3
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.Open, RoundOfType.Floor)
            Dim topWick As Decimal = ((candle.CandleWicks.Top + buffer) / candle.CandleRange) * 100
            Dim bottomWick As Decimal = ((candle.CandleWicks.Bottom + buffer) / candle.CandleRange) * 100
            If topWick >= 50 AndAlso bottomWick < 50 Then
                Dim buyLevel As Decimal = candle.High + buffer
                ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, buyLevel, candle, Trade.TradeExecutionDirection.Buy)
            ElseIf bottomWick >= 50 AndAlso topWick < 50 Then
                Dim sellLevel As Decimal = candle.Low - buffer
                ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, sellLevel, candle, Trade.TradeExecutionDirection.Sell)
            End If
        End If
        Return ret
    End Function
End Class
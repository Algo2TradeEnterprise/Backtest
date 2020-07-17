Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class SupertrendCutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
        Public StoplossPercentage As Decimal
    End Class
#End Region

    Private _supertrendPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _supertrendColorPayload As Dictionary(Of Date, Color) = Nothing

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

        Indicator.Supertrend.CalculateSupertrend(7, 3, _signalPayload, _supertrendPayload, _supertrendColorPayload)
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
            Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item3
                ElseIf lastExecutedOrder IsNot Nothing Then
                    If signal.Item3.PayloadDate <> lastExecutedOrder.SignalCandle.PayloadDate AndAlso
                        currentMinuteCandlePayload.PayloadDate > lastExecutedOrder.ExitTime Then
                        signalCandle = signal.Item3
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim entryPrice As Decimal = currentTick.Open
                Dim stoploss As Decimal = signal.Item2
                Dim slPoint As Decimal = ConvertFloorCeling(entryPrice * _userInputs.StoplossPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _supertrendPayload(signalCandle.PayloadDate)
                                }
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice + slPoint, entryPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Cash)

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _supertrendPayload(signalCandle.PayloadDate)
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            Dim supertrend As Decimal = _supertrendPayload(candle.PayloadDate)
            If _supertrendColorPayload(candle.PayloadDate) = Color.Green Then
                If candle.Low < supertrend AndAlso candle.Close > supertrend Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.Low, RoundOfType.Floor)
                    Dim diff As Decimal = (Math.Abs(currentTick.Open - (candle.Low - buffer)) / currentTick.Open) * 100
                    If diff <= _userInputs.StoplossPercentage Then
                        ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, candle.Low - buffer, candle, Trade.TradeExecutionDirection.Buy)
                    Else
                        Console.WriteLine(String.Format("Signal neglected. {0},{1}", _tradingSymbol, candle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss")))
                    End If
                End If
            ElseIf _supertrendColorPayload(candle.PayloadDate) = Color.Red Then
                If candle.High > supertrend AndAlso candle.Close < supertrend Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.High, RoundOfType.Floor)
                    Dim diff As Decimal = (Math.Abs(currentTick.Open - (candle.High + buffer)) / currentTick.Open) * 100
                    If diff <= _userInputs.StoplossPercentage Then
                        ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, candle.High + buffer, candle, Trade.TradeExecutionDirection.Sell)
                    Else
                        Console.WriteLine(String.Format("Signal neglected. {0},{1}", _tradingSymbol, candle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss")))
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class
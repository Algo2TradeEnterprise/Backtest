Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class ReversalHHLLBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MinimumTargetMultiplier As Decimal
        Public MaxLossPerTrade As Decimal
    End Class
#End Region

    Private ReadOnly _tradeStartTime As Date
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
        _tradeStartTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
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
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS)
                If lastExecutedTrade Is Nothing OrElse lastExecutedTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                    signalCandle = signal.Item3
                End If
            End If

            If signalCandle IsNot Nothing Then
                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandle.High + buffer
                    Dim stoploss As Decimal = signalCandle.Low - buffer
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, _userInputs.MaxLossPerTrade, _parentStrategy.StockType)

                    If currentTick.Open < entryPrice Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = signal.Item2,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
                    End If
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandle.Low - buffer
                    Dim stoploss As Decimal = signalCandle.High + buffer
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, _userInputs.MaxLossPerTrade, _parentStrategy.StockType)

                    If currentTick.Open > entryPrice Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = signal.Item2,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentTrade.SignalCandle.PayloadDate < currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing AndAlso candle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
            If candle.High > candle.PreviousCandlePayload.High AndAlso
                candle.PreviousCandlePayload.High > candle.PreviousCandlePayload.PreviousCandlePayload.High AndAlso
                candle.Low > candle.PreviousCandlePayload.Low AndAlso
                candle.PreviousCandlePayload.Low > candle.PreviousCandlePayload.PreviousCandlePayload.Low Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.Low, RoundOfType.Floor)
                Dim entry As Decimal = candle.Low - buffer
                Dim stoploss As Decimal = candle.High + buffer
                Dim target As Decimal = candle.PreviousCandlePayload.PreviousCandlePayload.Low
                If entry - target >= (stoploss - entry) * _userInputs.MinimumTargetMultiplier Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, target, candle, Trade.TradeExecutionDirection.Sell)
                Else
                    'Console.WriteLine(String.Format("Sell ignored Instrument:{0},Signal Candle:{1},Target:{2},Stoploss:{3}",
                    '                                candle.TradingSymbol, candle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss"), entry - target, stoploss - entry))
                End If
            ElseIf candle.High < candle.PreviousCandlePayload.High AndAlso
                candle.PreviousCandlePayload.High < candle.PreviousCandlePayload.PreviousCandlePayload.High AndAlso
                candle.Low < candle.PreviousCandlePayload.Low AndAlso
                candle.PreviousCandlePayload.Low < candle.PreviousCandlePayload.PreviousCandlePayload.Low Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.High, RoundOfType.Floor)
                Dim entry As Decimal = candle.High + buffer
                Dim stoploss As Decimal = candle.Low - buffer
                Dim target As Decimal = candle.PreviousCandlePayload.PreviousCandlePayload.High
                If target - entry >= (entry - stoploss) * _userInputs.MinimumTargetMultiplier Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, target, candle, Trade.TradeExecutionDirection.Buy)
                Else
                    'Console.WriteLine(String.Format("Buy ignored Instrument:{0},Signal Candle:{1},Target:{2},Stoploss:{3}",
                    '                                candle.TradingSymbol, candle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss"), target - entry, entry - stoploss))
                End If
            End If
        End If
        Return ret
    End Function
End Class
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation
Imports Utilities.Numbers

Public Class BollingerTouchStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
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
        Indicator.BollingerBands.CalculateBollingerBands(20, Payload.PayloadFields.Close, 2, _signalPayload, _bollingerHighPayload, _bollingerLowPayload, _smaPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandle, Trade.TypeOfTrade.MIS)
                If lastExecutedOrder IsNot Nothing Then
                    If lastExecutedOrder.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                        signalCandle = signal.Item3
                    End If
                Else
                        signalCandle = signal.Item3
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = signalCandle.High + buffer
                    Dim stoplossPrice As Decimal = signalCandle.Low - buffer
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoplossPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, _parentStrategy.StockType)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoplossPrice,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signalCandle.CandleRange,
                                    .Supporting3 = Math.Round(_atrPayload(signalCandle.PayloadDate), 4),
                                    .Supporting4 = Math.Round(_bollingerHighPayload(signalCandle.PayloadDate), 4),
                                    .Supporting5 = Math.Round(_bollingerLowPayload(signalCandle.PayloadDate), 4)
                                }
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = signalCandle.Low - buffer
                    Dim stoplossPrice As Decimal = signalCandle.High + buffer
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoplossPrice, entryPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, _parentStrategy.StockType)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Sell, _parentStrategy.StockType)

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoplossPrice,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signalCandle.CandleRange,
                                    .Supporting3 = Math.Round(_atrPayload(signalCandle.PayloadDate), 4),
                                    .Supporting4 = Math.Round(_bollingerHighPayload(signalCandle.PayloadDate), 4),
                                    .Supporting5 = Math.Round(_bollingerLowPayload(signalCandle.PayloadDate), 4)
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If (currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate) OrElse
                    (currentTrade.EntryDirection <> signal.Item4) Then
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim signalCandle As Payload = currentCandle.PreviousCandlePayload
            While signalCandle IsNot Nothing
                If signalCandle.PayloadDate.Date <> _tradingDate.Date Then Exit While
                If signalCandle.CandleRange <= _atrPayload(signalCandle.PayloadDate) * 2 / 3 Then
                    Dim takeTrade As Boolean = False
                    If signalCandle.High >= _bollingerHighPayload(signalCandle.PayloadDate) AndAlso signalCandle.Low <= _bollingerHighPayload(signalCandle.PayloadDate) Then
                        takeTrade = True
                    ElseIf signalCandle.High >= _bollingerLowPayload(signalCandle.PayloadDate) AndAlso signalCandle.Low <= _bollingerLowPayload(signalCandle.PayloadDate) Then
                        takeTrade = True
                    End If
                    If takeTrade Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                        Dim highPrice As Decimal = signalCandle.High + buffer
                        Dim lowPrice As Decimal = signalCandle.Low - buffer
                        Dim middlePrice As Decimal = (highPrice + lowPrice) / 2
                        Dim range As Decimal = highPrice - middlePrice
                        If currentTick.Open >= middlePrice + range * 60 / 100 Then
                            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, highPrice, signalCandle, Trade.TradeExecutionDirection.Buy)
                        ElseIf currentTick.Open <= middlePrice - range * 60 / 100 Then
                            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, lowPrice, signalCandle, Trade.TradeExecutionDirection.Sell)
                        End If
                        Exit While
                    End If
                End If

                signalCandle = signalCandle.PreviousCandlePayload
            End While
        End If
        Return ret
    End Function
End Class
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class BollingerCloseStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public ATRMultiplier As Decimal
        Public BollingerPeriod As Decimal
        Public StandardDeviation As Decimal
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
        Indicator.BollingerBands.CalculateBollingerBands(_userInputs.BollingerPeriod, Payload.PayloadFields.Close, _userInputs.StandardDeviation, _signalPayload, _bollingerHighPayload, _bollingerLowPayload, _smaPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso currentMinuteCandlePayload.PayloadDate >= _tradeStartTime Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item4
            End If

            If signalCandle IsNot Nothing Then
                'Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Open, RoundOfType.Floor)
                Dim buffer As Decimal = 1
                Dim atr As Decimal = _atrPayload(signalCandle.PayloadDate)
                If signal.Item5 = Trade.TradeExecutionDirection.Buy AndAlso
                    Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                    Dim entryPrice As Decimal = signalCandle.High + buffer
                    Dim quantity As Integer = signal.Item2

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = entryPrice - 100000000,
                                    .Target = entryPrice + 100000000,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item3,
                                    .Supporting3 = atr
                                }
                ElseIf signal.Item5 = Trade.TradeExecutionDirection.Sell AndAlso
                    Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                    Dim entryPrice As Decimal = signalCandle.Low - buffer
                    Dim quantity As Integer = signal.Item2

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = entryPrice + 100000000,
                                    .Target = entryPrice - 100000000,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item3,
                                    .Supporting3 = atr
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim bollingerHigh As Decimal = _bollingerHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
            Dim bollingerLow As Decimal = _bollingerLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso currentMinuteCandlePayload.PreviousCandlePayload.Close > bollingerHigh Then
                ret = New Tuple(Of Boolean, String)(True, "Reverse Signal")
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso currentMinuteCandlePayload.PreviousCandlePayload.Close < bollingerLow Then
                ret = New Tuple(Of Boolean, String)(True, "Reverse Signal")
            End If
        ElseIf currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signal As Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 AndAlso signal.Item4.PayloadDate <> currentTrade.SignalCandle.PayloadDate Then
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim direction As Tuple(Of Trade.TradeExecutionDirection, Payload) = GetSignalDirection(candle)
            If direction IsNot Nothing Then
                Dim atr As Decimal = _atrPayload(candle.PayloadDate) * _userInputs.ATRMultiplier
                Dim quantity As Integer = Me.LotSize
                Dim iteration As Integer = 1
                If direction.Item1 = Trade.TradeExecutionDirection.Buy Then
                    If candle.High < candle.PreviousCandlePayload.High Then
                        'Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.High, RoundOfType.Floor)
                        Dim buffer As Decimal = 1
                        Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(candle, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy)
                        If lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            If direction.Item2.PayloadDate > lastExecutedOrder.SignalCandle.PayloadDate AndAlso
                                lastExecutedOrder.EntryPrice - (candle.High + buffer) >= atr Then
                                iteration = Val(lastExecutedOrder.Supporting2) + 1
                                ret = New Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection)(True, quantity * iteration, iteration, candle, Trade.TradeExecutionDirection.Buy)
                            End If
                        Else
                            ret = New Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection)(True, quantity, iteration, candle, Trade.TradeExecutionDirection.Buy)
                        End If
                    End If
                ElseIf direction.Item1 = Trade.TradeExecutionDirection.Sell Then
                    If candle.Low > candle.PreviousCandlePayload.Low Then
                        'Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.Low, RoundOfType.Floor)
                        Dim buffer As Decimal = 1
                        Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(candle, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Sell)
                        If lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            If direction.Item2.PayloadDate > lastExecutedOrder.SignalCandle.PayloadDate AndAlso
                            (candle.Low - buffer) - lastExecutedOrder.EntryPrice >= atr Then
                                iteration = Val(lastExecutedOrder.Supporting2) + 1
                                ret = New Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection)(True, quantity * iteration, iteration, candle, Trade.TradeExecutionDirection.Sell)
                            End If
                        Else
                            ret = New Tuple(Of Boolean, Integer, Integer, Payload, Trade.TradeExecutionDirection)(True, quantity, iteration, candle, Trade.TradeExecutionDirection.Sell)
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetSignalDirection(ByVal candle As Payload) As Tuple(Of Trade.TradeExecutionDirection, Payload)
        Dim ret As Tuple(Of Trade.TradeExecutionDirection, Payload) = Nothing
        Dim lastExitOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(candle, _parentStrategy.TradeType)
        Dim startTime As Date = _tradingDate.Date
        If lastExitOrder IsNot Nothing Then
            Dim exitPayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExitOrder.ExitTime, _signalPayload))
            startTime = exitPayload.PreviousCandlePayload.PayloadDate
        End If
        For Each runningPayload In _signalPayload
            If runningPayload.Key >= startTime AndAlso runningPayload.Key <= candle.PayloadDate Then
                Dim bollingerHigh As Decimal = _bollingerHighPayload(runningPayload.Value.PayloadDate)
                Dim bollingerLow As Decimal = _bollingerLowPayload(runningPayload.Value.PayloadDate)
                If runningPayload.Value.Close > bollingerHigh Then
                    ret = New Tuple(Of Trade.TradeExecutionDirection, Payload)(Trade.TradeExecutionDirection.Sell, runningPayload.Value)
                ElseIf runningPayload.Value.Close < bollingerLow Then
                    ret = New Tuple(Of Trade.TradeExecutionDirection, Payload)(Trade.TradeExecutionDirection.Buy, runningPayload.Value)
                End If
            End If
        Next
        Return ret
    End Function
End Class
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class SqueezeBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
    End Class
#End Region

    Private _atrHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _atrLowPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _smaPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerLowPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _signalCandle As Payload = Nothing

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

        Indicator.ATRBands.CalculateATRBands(1, 50, Payload.PayloadFields.Close, _signalPayload, _atrHighPayload, _atrLowPayload)
        Indicator.BollingerBands.CalculateBollingerBands(20, Payload.PayloadFields.Close, 2, _signalPayload, _bollingerHighPayload, _bollingerLowPayload, _smaPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            Me.EligibleToTakeTrade AndAlso Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, Trade.TypeOfTrade.MIS) Then

            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item4
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)

                If signal.Item5 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = signal.Item2 + buffer
                    Dim stoploss As Decimal = signal.Item3
                    Dim iteration As Decimal = GetNumberOfTrade(signalCandle) + 1
                    Dim multiplier As Integer = Math.Pow(2, iteration - 1)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, Math.Abs(_userInputs.MaxLossPerTrade) * -1 * multiplier, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * multiplier, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)

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
                                .Supporting2 = iteration
                            }
                ElseIf signal.Item5 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = signal.Item2 - buffer
                    Dim stoploss As Decimal = signal.Item3
                    Dim iteration As Decimal = GetNumberOfTrade(signalCandle) + 1
                    Dim multiplier As Integer = Math.Pow(2, iteration - 1)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1 * multiplier, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * multiplier, Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Cash)

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
                                .Supporting2 = iteration
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
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryDirection <> signal.Item5 OrElse currentTrade.SignalCandle.PayloadDate <> signal.Item4.PayloadDate Then
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If _signalCandle Is Nothing Then
                If _bollingerHighPayload(candle.PreviousCandlePayload.PayloadDate) < _atrHighPayload(candle.PreviousCandlePayload.PayloadDate) AndAlso
                    _bollingerLowPayload(candle.PreviousCandlePayload.PayloadDate) > _atrLowPayload(candle.PreviousCandlePayload.PayloadDate) Then
                    If _bollingerHighPayload(candle.PayloadDate) > _atrHighPayload(candle.PayloadDate) OrElse
                        _bollingerLowPayload(candle.PayloadDate) < _atrLowPayload(candle.PayloadDate) Then
                        _signalCandle = candle
                    End If
                End If
            Else
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(candle, Trade.TypeOfTrade.MIS)
                If lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.SignalCandle.PayloadDate = _signalCandle.PayloadDate AndAlso
                    lastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Close AndAlso lastExecutedOrder.ExitCondition = Trade.TradeExitCondition.Target Then
                    _signalCandle = Nothing
                End If
            End If

            If _signalCandle IsNot Nothing Then
                Dim buyEntryPrice As Decimal = ConvertFloorCeling(_atrHighPayload(_signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                Dim sellEntryPrice As Decimal = ConvertFloorCeling(_atrLowPayload(_signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
                Dim middlePrice As Decimal = (buyEntryPrice + sellEntryPrice) / 2
                Dim range As Decimal = buyEntryPrice - middlePrice
                If currentTick.Open >= middlePrice + range * 60 / 100 Then
                    Dim stoploss As Decimal = ConvertFloorCeling(middlePrice, _parentStrategy.TickSize, RoundOfType.Floor)
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection)(True, buyEntryPrice, stoploss, _signalCandle, Trade.TradeExecutionDirection.Buy)
                ElseIf currentTick.Open <= middlePrice - range * 60 / 100 Then
                    Dim stoploss As Decimal = ConvertFloorCeling(middlePrice, _parentStrategy.TickSize, RoundOfType.Celing)
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection)(True, sellEntryPrice, stoploss, _signalCandle, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetNumberOfTrade(ByVal signalCandle As Payload) As Integer
        Dim ret As Integer = 0
        Dim closedOrders As List(Of Trade) = _parentStrategy.GetSpecificTrades(signalCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Close)
        Dim inprogressOrders As List(Of Trade) = _parentStrategy.GetSpecificTrades(signalCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        Dim totalExecutedOrders As List(Of Trade) = New List(Of Trade)
        If closedOrders IsNot Nothing Then totalExecutedOrders.AddRange(closedOrders)
        If inprogressOrders IsNot Nothing Then totalExecutedOrders.AddRange(inprogressOrders)
        If totalExecutedOrders IsNot Nothing AndAlso totalExecutedOrders.Count > 0 Then
            Dim signalTrades As List(Of Trade) = totalExecutedOrders.FindAll(Function(x)
                                                                                 Return x.SignalCandle.PayloadDate = signalCandle.PayloadDate
                                                                             End Function)
            If signalTrades IsNot Nothing Then ret = signalTrades.Count
        End If
        Return ret
    End Function
End Class
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HKReversalMartingaleStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxProfitPerTrade As Decimal
        Public MaxLossPerTrade As Decimal
    End Class
#End Region

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _emaPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _slReamrk As String = ""
    Private _slPoint As Decimal = Decimal.MinValue
    Private _targetPoint As Decimal = Decimal.MinValue
    Private _quantity As Integer = Integer.MinValue

    Private ReadOnly _direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Private ReadOnly _userInputs As StrategyRuleEntities

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal direction As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        If direction > 0 Then
            _direction = Trade.TradeExecutionDirection.Buy
        ElseIf direction < 0 Then
            _direction = Trade.TradeExecutionDirection.Sell
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
        Indicator.ATR.CalculateATR(14, _hkPayload, _atrPayload)
        Indicator.EMA.CalculateEMA(13, Payload.PayloadFields.Close, _hkPayload, _emaPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = Nothing
            Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item3

                    Dim highestATR As Decimal = ConvertFloorCeling(GetHighestATR(signalCandle), _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim candleRange As Decimal = ConvertFloorCeling(signalCandle.CandleRange, _parentStrategy.TickSize, RoundOfType.Celing)
                    If candleRange < highestATR / 2 Then
                        _slPoint = ConvertFloorCeling(highestATR / 2, _parentStrategy.TickSize, RoundOfType.Celing)
                        _slReamrk = "Half ATR"
                    Else
                        If candleRange < highestATR Then
                            _slPoint = candleRange
                            _slReamrk = "Candle Range"
                        Else
                            _slPoint = highestATR
                            _slReamrk = "Highest ATR"
                        End If
                    End If
                    _quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signal.Item2, signal.Item2 - _slPoint, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                    _targetPoint = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, signal.Item2, _quantity, _userInputs.MaxProfitPerTrade, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - signal.Item2
                ElseIf lastExecutedOrder IsNot Nothing Then
                    signalCandle = signal.Item3
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim entryPrice As Decimal = signal.Item2
                Dim slPoint As Decimal = _slPoint
                Dim targetPoint As Decimal = _targetPoint
                Dim quantity As Integer = _quantity
                If lastExecutedOrder IsNot Nothing Then
                    quantity = lastExecutedOrder.Quantity * 2
                End If

                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _slReamrk,
                                    .Supporting3 = ConvertFloorCeling(GetHighestATR(signalCandle), _parentStrategy.TickSize, RoundOfType.Celing),
                                    .Supporting4 = ConvertFloorCeling(signalCandle.CandleRange, _parentStrategy.TickSize, RoundOfType.Celing),
                                    .Supporting5 = _parentStrategy.StockNumberOfTrades(_tradingDate, _tradingSymbol) + 1
                                }
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _slReamrk,
                                    .Supporting3 = ConvertFloorCeling(GetHighestATR(signalCandle), _parentStrategy.TickSize, RoundOfType.Celing),
                                    .Supporting4 = ConvertFloorCeling(signalCandle.CandleRange, _parentStrategy.TickSize, RoundOfType.Celing),
                                    .Supporting5 = _parentStrategy.StockNumberOfTrades(_tradingDate, _tradingSymbol) + 1
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
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
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
            Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentCandle, _parentStrategy.TradeType)
            If lastExecutedOrder Is Nothing OrElse currentCandle.PayloadDate > lastExecutedOrder.ExitTime Then
                Dim hkCandle As Payload = _hkPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                Dim ema As Decimal = _emaPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                If _direction = Trade.TradeExecutionDirection.Buy AndAlso hkCandle.Open < ema AndAlso Math.Round(hkCandle.High, 4) = Math.Round(hkCandle.Open, 4) Then
                    Dim buyLevel As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                    If buyLevel = hkCandle.High Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(buyLevel, RoundOfType.Floor)
                        buyLevel = buyLevel + buffer
                    End If
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, buyLevel, hkCandle, Trade.TradeExecutionDirection.Buy)
                ElseIf _direction = Trade.TradeExecutionDirection.Sell AndAlso hkCandle.Open > ema AndAlso Math.Round(hkCandle.Low, 4) = Math.Round(hkCandle.Open, 4) Then
                    Dim sellLevel As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                    If sellLevel = hkCandle.Low Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(sellLevel, RoundOfType.Floor)
                        sellLevel = sellLevel - buffer
                    End If
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, sellLevel, hkCandle, Trade.TradeExecutionDirection.Sell)
                End If
            Else
                ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, lastExecutedOrder.EntryPrice, lastExecutedOrder.SignalCandle, lastExecutedOrder.EntryDirection)
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 AndAlso signalCandle IsNot Nothing Then
            Dim todayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                 If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                                                     Return x.Value
                                                                 Else
                                                                     Return Decimal.MinValue
                                                                 End If
                                                             End Function)
            If todayHighestATR <> Decimal.MinValue Then
                ret = todayHighestATR
            End If
        End If
        Return ret
    End Function
End Class
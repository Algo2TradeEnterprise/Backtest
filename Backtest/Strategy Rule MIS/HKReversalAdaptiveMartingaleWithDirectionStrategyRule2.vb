Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HKReversalAdaptiveMartingaleWithDirectionStrategyRule2
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
    Private _vwapPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _emaPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _previousTradingDay As Date

    Private _slPoint As Decimal

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

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)

        Indicator.ATR.CalculateATR(14, _hkPayload, _atrPayload)
        Indicator.VWAP.CalculateVWAP(_hkPayload, _vwapPayload)
        Indicator.EMA.CalculateEMA(20, Payload.PayloadFields.Close, _hkPayload, _emaPayload)
        _previousTradingDay = _parentStrategy.Cmn.GetPreviousTradingDay(Common.DataBaseTable.Intraday_Cash, _tradingSymbol, _tradingDate)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandle, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandle, _parentStrategy.TradeType)
                If lastExecutedOrder IsNot Nothing Then
                    signalCandle = signal.Item3
                Else
                    signalCandle = signal.Item3
                    _slPoint = ConvertFloorCeling(GetAverageHighestATR(signalCandle), _parentStrategy.TickSize, RoundOfType.Celing)
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim entryPrice As Decimal = signal.Item2
                Dim slPoint As Decimal = _slPoint
                Dim iteration As Integer = _parentStrategy.StockNumberOfTrades(_tradingDate, _tradingSymbol) + 1
                Dim lossPL As Decimal = _userInputs.MaxLossPerTrade + _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                If lossPL < 0 Then
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, lossPL, Trade.TypeOfStock.Cash)
                    Dim profitPL As Decimal = _userInputs.MaxProfitPerTrade - _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                    Dim targetPoint As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, profitPL, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - entryPrice

                    If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item5,
                                    .Supporting3 = iteration
                                }
                    ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item5,
                                    .Supporting3 = iteration
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
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentMinuteCandle.PreviousCandlePayload.Close < currentTrade.SignalCandle.Low Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentMinuteCandle.PreviousCandlePayload.Close > currentTrade.SignalCandle.High Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
            If ret Is Nothing Then
                Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandle, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    If currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim direction As Trade.TradeExecutionDirection = GetEntryDirection(currentCandle, currentTick)
            If direction = Trade.TradeExecutionDirection.Buy Then
                Dim signalCandle As Payload = Nothing
                For Each runningPayload In _hkPayload.OrderByDescending(Function(x)
                                                                            Return x.Key
                                                                        End Function)
                    If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key >= _tradeStartTime AndAlso
                        runningPayload.Key < currentCandle.PayloadDate Then
                        If Math.Round(runningPayload.Value.Open, 2) = Math.Round(runningPayload.Value.High, 2) Then
                            signalCandle = runningPayload.Value
                            Exit For
                        End If
                    End If
                Next
                If signalCandle IsNot Nothing Then
                    Dim buyLevel As Decimal = ConvertFloorCeling(signalCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                    If buyLevel = Math.Round(signalCandle.High, 2) Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(buyLevel, RoundOfType.Floor)
                        buyLevel = buyLevel + buffer
                    End If
                    If Not IsSignalTriggered(buyLevel, Trade.TradeExecutionDirection.Buy, signalCandle.PayloadDate, currentCandle.PayloadDate) Then
                        Dim candleClosedBelowSignalCandle As Boolean = False
                        For Each runningPayload In _hkPayload
                            If runningPayload.Key >= signalCandle.PayloadDate AndAlso runningPayload.Key < currentCandle.PayloadDate Then
                                If runningPayload.Value.Close < signalCandle.Low Then
                                    candleClosedBelowSignalCandle = True
                                    Exit For
                                End If
                            End If
                        Next
                        If Not candleClosedBelowSignalCandle Then
                            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, buyLevel, signalCandle, Trade.TradeExecutionDirection.Buy, "")
                        End If
                    End If
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                Dim signalCandle As Payload = Nothing
                For Each runningPayload In _hkPayload.OrderByDescending(Function(x)
                                                                            Return x.Key
                                                                        End Function)
                    If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key >= _tradeStartTime AndAlso
                        runningPayload.Key < currentCandle.PayloadDate Then
                        If Math.Round(runningPayload.Value.Open, 2) = Math.Round(runningPayload.Value.Low, 2) Then
                            signalCandle = runningPayload.Value
                            Exit For
                        End If
                    End If
                Next
                If signalCandle IsNot Nothing Then
                    Dim sellLevel As Decimal = ConvertFloorCeling(signalCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                    If sellLevel = Math.Round(signalCandle.Low, 2) Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(sellLevel, RoundOfType.Floor)
                        sellLevel = sellLevel - buffer
                    End If
                    If Not IsSignalTriggered(sellLevel, Trade.TradeExecutionDirection.Sell, signalCandle.PayloadDate, currentCandle.PayloadDate) Then
                        Dim candleClosedBelowSignalCandle As Boolean = False
                        For Each runningPayload In _hkPayload
                            If runningPayload.Key >= signalCandle.PayloadDate AndAlso runningPayload.Key < currentCandle.PayloadDate Then
                                If runningPayload.Value.Close > signalCandle.High Then
                                    candleClosedBelowSignalCandle = True
                                    Exit For
                                End If
                            End If
                        Next
                        If Not candleClosedBelowSignalCandle Then
                            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, sellLevel, signalCandle, Trade.TradeExecutionDirection.Sell, "")
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetEntryDirection(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Trade.TradeExecutionDirection
        Dim ret As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            If _emaPayload(currentCandle.PreviousCandlePayload.PayloadDate) > _vwapPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                ret = Trade.TradeExecutionDirection.Buy
            ElseIf _emaPayload(currentCandle.PreviousCandlePayload.PayloadDate) < _vwapPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                ret = Trade.TradeExecutionDirection.Sell
            End If
        End If
        Return ret
    End Function

    Private Function GetAverageHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 AndAlso signalCandle IsNot Nothing Then
            Dim todayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                 If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                                                     Return x.Value
                                                                 Else
                                                                     Return Decimal.MinValue
                                                                 End If
                                                             End Function)

            Dim previousDayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                       If x.Key.Date = _previousTradingDay.Date Then
                                                                           Return x.Value
                                                                       Else
                                                                           Return Decimal.MinValue
                                                                       End If
                                                                   End Function)

            If todayHighestATR <> Decimal.MinValue AndAlso previousDayHighestATR <> Decimal.MinValue Then
                ret = (todayHighestATR + previousDayHighestATR) / 2
            End If
        End If
        Return ret
    End Function
End Class
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HKReversalAdaptiveMartingaleWithDirectionStrategyRule1
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxProfitPerTrade As Decimal
        Public MaxLossPerTrade As Decimal
        Public MoveStoplossToBreakoutCandleHL As Boolean
        Public MaxCapitalPerTrade As Decimal
        Public MinCapitalPerTrade As Decimal
    End Class
#End Region

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _vwapPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _previousTradingDay As Date

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
        Indicator.FractalBands.CalculateFractal(_hkPayload, _fractalHighPayload, _fractalLowPayload)
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
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandle, _parentStrategy.TradeType)
                If lastExecutedOrder IsNot Nothing Then
                    Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedOrder.ExitTime, _signalPayload))
                    If signal.Item5.PayloadDate >= lastExecutedOrder.ExitTime Then
                        signalCandle = signal.Item5
                    End If
                Else
                    signalCandle = signal.Item5
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim entryPrice As Decimal = signal.Item2
                Dim stoploss As Decimal = signal.Item3
                Dim slPoint As Decimal = signal.Item4
                Dim iteration As Integer = _parentStrategy.StockNumberOfTrades(_tradingDate, _tradingSymbol) + 1
                Dim lossPL As Decimal = _userInputs.MaxLossPerTrade + _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                If lossPL < 0 Then
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, lossPL, Trade.TypeOfStock.Cash)
                    Dim profitPL As Decimal = _userInputs.MaxProfitPerTrade - _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                    Dim targetPoint As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, profitPL, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - entryPrice

                    If signal.Item6 = Trade.TradeExecutionDirection.Buy Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item7,
                                    .Supporting3 = iteration
                                }
                    ElseIf signal.Item6 = Trade.TradeExecutionDirection.Sell Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item7,
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
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signal.Item5.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim remarks As String = Nothing
            If _userInputs.MoveStoplossToBreakoutCandleHL Then
                Dim breakoutCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTrade.EntryTime, _signalPayload))
                If currentTick.PayloadDate >= breakoutCandle.PayloadDate.AddMinutes(_parentStrategy.SignalTimeFrame) Then
                    If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        triggerPrice = ConvertFloorCeling(breakoutCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                        If triggerPrice = breakoutCandle.Low Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(triggerPrice, RoundOfType.Floor)
                            triggerPrice = breakoutCandle.Low - buffer
                        End If
                        remarks = currentTrade.EntryPrice - triggerPrice
                        If currentTick.Open <= triggerPrice Then triggerPrice = Decimal.MinValue
                    ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        triggerPrice = ConvertFloorCeling(breakoutCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                        If triggerPrice = breakoutCandle.High Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(triggerPrice, RoundOfType.Floor)
                            triggerPrice = breakoutCandle.High + buffer
                        End If
                        remarks = triggerPrice - currentTrade.EntryPrice
                        If currentTick.Open >= triggerPrice Then triggerPrice = Decimal.MinValue
                    End If
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso currentTrade.PotentialStopLoss <> triggerPrice Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, remarks)
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayload(currentCandle.PreviousCandlePayload.PayloadDate)
            Dim direction As Trade.TradeExecutionDirection = GetEntryDirection(currentCandle, currentTick)
            If direction = Trade.TradeExecutionDirection.Buy AndAlso Math.Round(hkCandle.High, 2) = Math.Round(hkCandle.Open, 2) Then
                Dim buyLevel As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                If buyLevel = hkCandle.High Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(buyLevel, RoundOfType.Floor)
                    buyLevel = buyLevel + buffer
                End If
                Dim sellLevel As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                If sellLevel = hkCandle.Low Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(sellLevel, RoundOfType.Floor)
                    sellLevel = sellLevel - buffer
                End If
                Dim qunatity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, buyLevel, sellLevel, _userInputs.MaxLossPerTrade, _parentStrategy.StockType)
                Dim capital As Decimal = buyLevel * qunatity / _parentStrategy.MarginMultiplier
                If capital >= _userInputs.MinCapitalPerTrade Then
                    If capital <= _userInputs.MaxCapitalPerTrade Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, buyLevel, sellLevel, buyLevel - sellLevel, hkCandle, Trade.TradeExecutionDirection.Buy, "Normal")
                    Else
                        Dim slPoint As Decimal = buyLevel - sellLevel
                        For sl As Decimal = sellLevel To 0 Step -1 * _parentStrategy.TickSize
                            Dim modifiedQunatity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, buyLevel, sl, _userInputs.MaxLossPerTrade, _parentStrategy.StockType)
                            Dim modifiedCapital As Decimal = buyLevel * modifiedQunatity / _parentStrategy.MarginMultiplier
                            If modifiedCapital <= _userInputs.MaxCapitalPerTrade Then
                                slPoint = buyLevel - sl
                                Exit For
                            End If
                        Next
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, buyLevel, sellLevel, slPoint, hkCandle, Trade.TradeExecutionDirection.Buy, "Adjusted")
                    End If
                Else
                    'Console.WriteLine(String.Format("Trade neglected for higher capital. Signal Candle:{0}, Direction:{1}, Capital:{2}, Trading Symbol:{3}",
                    '                                hkCandle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss"), "BUY", Math.Round(capital, 2), _tradingSymbol))
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell AndAlso Math.Round(hkCandle.Low, 2) = Math.Round(hkCandle.Open, 2) Then
                Dim sellLevel As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                If sellLevel = hkCandle.Low Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(sellLevel, RoundOfType.Floor)
                    sellLevel = sellLevel - buffer
                End If
                Dim buyLevel As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                If buyLevel = hkCandle.High Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(buyLevel, RoundOfType.Floor)
                    buyLevel = buyLevel + buffer
                End If
                Dim qunatity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, buyLevel, sellLevel, _userInputs.MaxLossPerTrade, _parentStrategy.StockType)
                Dim capital As Decimal = sellLevel * qunatity / _parentStrategy.MarginMultiplier
                If capital >= _userInputs.MinCapitalPerTrade Then
                    If capital <= _userInputs.MaxCapitalPerTrade Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, sellLevel, buyLevel, buyLevel - sellLevel, hkCandle, Trade.TradeExecutionDirection.Sell, "Normal")
                    Else
                        Dim slPoint As Decimal = buyLevel - sellLevel
                        For sl As Decimal = buyLevel To Decimal.MaxValue Step 1 * _parentStrategy.TickSize
                            Dim modifiedQunatity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, sl, sellLevel, _userInputs.MaxLossPerTrade, _parentStrategy.StockType)
                            Dim modifiedCapital As Decimal = sellLevel * modifiedQunatity / _parentStrategy.MarginMultiplier
                            If modifiedCapital <= _userInputs.MaxCapitalPerTrade Then
                                slPoint = sl - sellLevel
                                Exit For
                            End If
                        Next
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, sellLevel, buyLevel, slPoint, hkCandle, Trade.TradeExecutionDirection.Sell, "Adjusted")
                    End If
                Else
                    'Console.WriteLine(String.Format("Trade neglected for higher capital. Signal Candle:{0}, Direction:{1}, Capital:{2}, Trading Symbol:{3}",
                    '                                hkCandle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss"), "SELL", Math.Round(capital, 2), _tradingSymbol))
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetEntryDirection(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Trade.TradeExecutionDirection
        Dim ret As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
        For Each runningPayload In _hkPayload.OrderByDescending(Function(x)
                                                                    Return x.Key
                                                                End Function)
            If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key < currentCandle.PayloadDate Then
                If _fractalHighPayload(runningPayload.Key) > _vwapPayload(runningPayload.Key) AndAlso
                    _fractalLowPayload(runningPayload.Key) > _vwapPayload(runningPayload.Key) Then
                    ret = Trade.TradeExecutionDirection.Buy
                    Exit For
                ElseIf _fractalHighPayload(runningPayload.Key) < _vwapPayload(runningPayload.Key) AndAlso
                    _fractalLowPayload(runningPayload.Key) < _vwapPayload(runningPayload.Key) Then
                    ret = Trade.TradeExecutionDirection.Sell
                    Exit For
                End If
            End If
        Next
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
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HKSlabBasedStrategyRule
    Inherits StrategyRule

    Private _hkPayloads As Dictionary(Of Date, Payload)

    Private _lastBuySignal As Payload = Nothing
    Private _lastSellSignal As Payload = Nothing

    Private ReadOnly _slab As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal slab As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _slab = slab
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayloads)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signalCandle As Payload = Nothing
                Dim signal As Tuple(Of Boolean, Decimal, Payload, String) = GetSignalCandle(currentMinuteCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    signalCandle = signal.Item3
                End If
                If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2
                    Dim slPrice As Decimal = signalCandle.Low - buffer
                    Dim quantity As Decimal = Me.LotSize
                    Dim target As Decimal = _slab * 3

                    If currentTick.Open < entryPrice Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = slPrice,
                                    .Target = .EntryPrice + target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item4
                                }
                    End If
                End If
            End If
            If parameter Is Nothing AndAlso Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signalCandle As Payload = Nothing
                Dim signal As Tuple(Of Boolean, Decimal, Payload, String) = GetSignalCandle(currentMinuteCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    signalCandle = signal.Item3
                End If

                If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2
                    Dim slPrice As Decimal = signalCandle.High + buffer
                    Dim quantity As Decimal = Me.LotSize
                    Dim target As Decimal = _slab * 3

                    If currentTick.Open > entryPrice Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = slPrice,
                                    .Target = .EntryPrice - target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item4
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS) Then
            ret = New Tuple(Of Boolean, String)(True, "One trade target reached")
        End If
        If ret Is Nothing Then
            If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
                Dim signal As Tuple(Of Boolean, Decimal, Payload, String) = GetSignalCandle(currentMinuteCandlePayload, currentTick, currentTrade.EntryDirection)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    If signal.Item3.PayloadDate <> currentTrade.SignalCandle.PayloadDate Then
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim remark As String = Nothing
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim signal As Tuple(Of Boolean, Decimal, Payload, String) = GetSignalCandle(currentMinuteCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim signalTrade As Trade = GetLastOrderFromSignalCandle(Trade.TradeExecutionDirection.Sell, signal.Item3)
                    If signalTrade Is Nothing Then
                        If signal.Item3.Low < currentTrade.EntryPrice Then
                            triggerPrice = signal.Item3.Low
                            remark = "Opposite signal low"
                        End If
                    End If
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim signal As Tuple(Of Boolean, Decimal, Payload, String) = GetSignalCandle(currentMinuteCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim signalTrade As Trade = GetLastOrderFromSignalCandle(Trade.TradeExecutionDirection.Buy, signal.Item3)
                    If signalTrade Is Nothing Then
                        If signal.Item3.High > currentTrade.EntryPrice Then
                            triggerPrice = signal.Item3.Low
                            remark = "Opposite signal high"
                        End If
                    End If
                End If
            End If
            If triggerPrice = Decimal.MinValue Then
                Dim entryCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTrade.EntryTime, _signalPayload))
                Dim previousHKCandle As Payload = _hkPayloads(currentMinuteCandlePayload.PayloadDate).PreviousCandlePayload
                If previousHKCandle.PayloadDate >= entryCandlePayload.PayloadDate Then
                    If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        Dim buyLevel As Decimal = GetSlabBasedLevel(currentTrade.SignalCandle.High, Trade.TradeExecutionDirection.Buy)
                        If currentTrade.SLRemark.ToUpper <> "NORMAL SL" AndAlso previousHKCandle.Close < buyLevel Then
                            triggerPrice = GetSlabBasedLevel(previousHKCandle.Low, Trade.TradeExecutionDirection.Sell)
                            remark = "Force SL"
                        Else
                            triggerPrice = currentTrade.SignalCandle.Low - currentTrade.StoplossBuffer
                            remark = "Normal SL"
                        End If
                    ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        Dim sellLevel As Decimal = GetSlabBasedLevel(currentTrade.SignalCandle.Low, Trade.TradeExecutionDirection.Sell)
                        If currentTrade.SLRemark.ToUpper <> "NORMAL SL" AndAlso previousHKCandle.Close > sellLevel Then
                            triggerPrice = GetSlabBasedLevel(previousHKCandle.High, Trade.TradeExecutionDirection.Buy)
                            remark = "Force SL"
                        Else
                            triggerPrice = currentTrade.SignalCandle.High + currentTrade.StoplossBuffer
                            remark = "Normal SL"
                        End If
                    End If
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, remark)
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

    Private Function GetSlabBasedLevel(ByVal price As Decimal, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If direction = Trade.TradeExecutionDirection.Buy Then
            ret = Math.Ceiling(price / _slab) * _slab
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            ret = Math.Floor(price / _slab) * _slab
        End If
        Return ret
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Decimal, Payload, String)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, String) = Nothing
        Dim buySignal As Payload = GetBuySignalCandle(candle.PayloadDate)
        If buySignal IsNot Nothing Then _lastBuySignal = buySignal
        Dim sellSignal As Payload = GetSellSignalCandle(candle.PayloadDate)
        If sellSignal IsNot Nothing Then _lastSellSignal = sellSignal
        Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(candle, Trade.TypeOfTrade.MIS)
        If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.SLRemark.ToUpper = "FORCE SL" AndAlso
            lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If lastExecutedTrade.EntryDirection <> direction Then
                If direction = Trade.TradeExecutionDirection.Buy Then
                    Dim lastTrade As Trade = GetLastOrder(Trade.TradeExecutionDirection.Buy, candle)
                    If lastTrade IsNot Nothing AndAlso lastTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                        Dim buyPrice As Decimal = GetSlabBasedLevel(lastTrade.SignalCandle.High, Trade.TradeExecutionDirection.Buy)
                        ret = New Tuple(Of Boolean, Decimal, Payload, String)(True, buyPrice, lastTrade.SignalCandle, "Force Entry")
                    End If
                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                    Dim lastTrade As Trade = GetLastOrder(Trade.TradeExecutionDirection.Sell, candle)
                    If lastTrade IsNot Nothing AndAlso lastTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                        Dim sellPrice As Decimal = GetSlabBasedLevel(lastTrade.SignalCandle.Low, Trade.TradeExecutionDirection.Sell)
                        ret = New Tuple(Of Boolean, Decimal, Payload, String)(True, sellPrice, lastTrade.SignalCandle, "Force Entry")
                    End If
                End If
            End If
        End If
        If ret Is Nothing Then
            If direction = Trade.TradeExecutionDirection.Buy Then
                If _lastBuySignal IsNot Nothing Then
                    Dim lastTrade As Trade = GetLastOrder(Trade.TradeExecutionDirection.Buy, candle)
                    If lastTrade Is Nothing OrElse lastTrade.SignalCandle.PayloadDate <> _lastBuySignal.PayloadDate Then
                        Dim buyPrice As Decimal = GetSlabBasedLevel(_lastBuySignal.High, Trade.TradeExecutionDirection.Buy)
                        ret = New Tuple(Of Boolean, Decimal, Payload, String)(True, buyPrice, _lastBuySignal, "Normal Entry")
                    End If
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                If _lastSellSignal IsNot Nothing Then
                    Dim lastTrade As Trade = GetLastOrder(Trade.TradeExecutionDirection.Sell, candle)
                    If lastTrade Is Nothing OrElse lastTrade.SignalCandle.PayloadDate <> _lastSellSignal.PayloadDate Then
                        Dim sellPrice As Decimal = GetSlabBasedLevel(_lastSellSignal.Low, Trade.TradeExecutionDirection.Sell)
                        ret = New Tuple(Of Boolean, Decimal, Payload, String)(True, sellPrice, _lastSellSignal, "Normal Entry")
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetBuySignalCandle(ByVal beforeThisTime As Date) As Payload
        Dim ret As Payload = Nothing
        If _lastBuySignal IsNot Nothing Then
            Dim signalTrade As Trade = GetLastOrderFromSignalCandle(Trade.TradeExecutionDirection.Buy, _lastBuySignal)
            If signalTrade IsNot Nothing AndAlso signalTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                _lastBuySignal = Nothing
            End If
        End If
        For Each runningPayload In _hkPayloads.OrderByDescending(Function(x)
                                                                     Return x.Key
                                                                 End Function)
            If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key < beforeThisTime Then
                If runningPayload.Value.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                    Dim level As Decimal = GetSlabBasedLevel(runningPayload.Value.High, Trade.TradeExecutionDirection.Sell)
                    If runningPayload.Value.Open > level AndAlso runningPayload.Value.Close < level Then
                        If _lastBuySignal Is Nothing OrElse
                            (_lastBuySignal.High > runningPayload.Value.High AndAlso runningPayload.Value.PayloadDate > _lastBuySignal.PayloadDate) Then
                            ret = runningPayload.Value
                            Exit For
                        End If
                    End If
                End If
            ElseIf runningPayload.Key.Date <> _tradingDate.Date Then
                Exit For
            End If
        Next
        Return ret
    End Function

    Private Function GetSellSignalCandle(ByVal beforeThisTime As Date) As Payload
        Dim ret As Payload = Nothing
        If _lastSellSignal IsNot Nothing Then
            Dim signalTrade As Trade = GetLastOrderFromSignalCandle(Trade.TradeExecutionDirection.Sell, _lastSellSignal)
            If signalTrade IsNot Nothing AndAlso signalTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                _lastSellSignal = Nothing
            End If
        End If
        For Each runningPayload In _hkPayloads.OrderByDescending(Function(x)
                                                                     Return x.Key
                                                                 End Function)
            If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key < beforeThisTime Then
                If runningPayload.Value.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                    Dim level As Decimal = GetSlabBasedLevel(runningPayload.Value.Low, Trade.TradeExecutionDirection.Buy)
                    If runningPayload.Value.Open < level AndAlso runningPayload.Value.Close > level Then
                        If _lastSellSignal Is Nothing OrElse
                            (_lastSellSignal.Low < runningPayload.Value.Low AndAlso runningPayload.Value.PayloadDate > _lastSellSignal.PayloadDate) Then
                            ret = runningPayload.Value
                            Exit For
                        End If
                    End If
                End If
            ElseIf runningPayload.Key.Date <> _tradingDate.Date Then
                Exit For
            End If
        Next
        Return ret
    End Function

    Private Function GetLastOrder(ByVal direction As Trade.TradeExecutionDirection, ByVal candle As Payload) As Trade
        Dim ret As Trade = Nothing
        Dim cancelTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Cancel)
        Dim closeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Close)
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        Dim openTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Open)
        Dim allTrades As List(Of Trade) = New List(Of Trade)
        If cancelTrades IsNot Nothing AndAlso cancelTrades.Count > 0 Then allTrades.AddRange(cancelTrades)
        If closeTrades IsNot Nothing AndAlso closeTrades.Count > 0 Then allTrades.AddRange(closeTrades)
        If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then allTrades.AddRange(inprogressTrades)
        If openTrades IsNot Nothing AndAlso openTrades.Count > 0 Then allTrades.AddRange(openTrades)
        If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
            Dim specificDirectionTrades As List(Of Trade) = allTrades.FindAll(Function(x)
                                                                                  Return x.EntryDirection = direction
                                                                              End Function)
            If specificDirectionTrades IsNot Nothing AndAlso specificDirectionTrades.Count > 0 Then
                ret = specificDirectionTrades.OrderBy(Function(x)
                                                          Return x.EntryTime
                                                      End Function).LastOrDefault
            End If
        End If
        Return ret
    End Function

    Private Function GetLastOrderFromSignalCandle(ByVal direction As Trade.TradeExecutionDirection, ByVal signalCandle As Payload) As Trade
        Dim ret As Trade = Nothing
        Dim cancelTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(signalCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Cancel)
        Dim closeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(signalCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Close)
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(signalCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        Dim openTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(signalCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Open)
        Dim allTrades As List(Of Trade) = New List(Of Trade)
        If cancelTrades IsNot Nothing AndAlso cancelTrades.Count > 0 Then allTrades.AddRange(cancelTrades)
        If closeTrades IsNot Nothing AndAlso closeTrades.Count > 0 Then allTrades.AddRange(closeTrades)
        If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then allTrades.AddRange(inprogressTrades)
        If openTrades IsNot Nothing AndAlso openTrades.Count > 0 Then allTrades.AddRange(openTrades)
        If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
            Dim specificDirectionTrades As List(Of Trade) = allTrades.FindAll(Function(x)
                                                                                  Return x.EntryDirection = direction
                                                                              End Function)
            If specificDirectionTrades IsNot Nothing AndAlso specificDirectionTrades.Count > 0 Then
                Dim signalTrades As List(Of Trade) = specificDirectionTrades.FindAll(Function(x)
                                                                                         Return x.SignalCandle.PayloadDate = signalCandle.PayloadDate
                                                                                     End Function)

                If signalTrades IsNot Nothing AndAlso signalTrades.Count > 0 Then
                    ret = signalTrades.OrderBy(Function(x)
                                                   Return x.EntryTime
                                               End Function).LastOrDefault
                End If
            End If
        End If
        Return ret
    End Function
End Class
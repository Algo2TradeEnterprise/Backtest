Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HighLowSlabLevelBasedStrategyRule
    Inherits StrategyRule

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
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signalCandle As Payload = Nothing
                Dim signal As Tuple(Of Boolean, Decimal, Payload, String) = GetSignalCandle(currentMinuteCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    signalCandle = signal.Item3
                End If
                If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2 + buffer
                    Dim slPrice As Decimal = signal.Item2 - _slab - buffer
                    Dim quantity As Decimal = Me.LotSize
                    Dim target As Decimal = _slab * 10

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
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2 - buffer
                    Dim slPrice As Decimal = signal.Item2 + _slab + buffer
                    Dim quantity As Decimal = Me.LotSize
                    Dim target As Decimal = _slab * 10

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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signal As Tuple(Of Boolean, Decimal, Payload, String) = GetSignalCandle(currentMinuteCandlePayload, currentTick, currentTrade.EntryDirection)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If signal.Item3.PayloadDate <> currentTrade.SignalCandle.PayloadDate Then
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
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim remark As String = Nothing
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim lastSellOrder As Trade = GetLastOrder(Trade.TradeExecutionDirection.Sell, currentMinuteCandlePayload)
                If lastSellOrder IsNot Nothing AndAlso lastSellOrder.TradeCurrentStatus <> Trade.TradeExecutionStatus.Close AndAlso
                    lastSellOrder.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel AndAlso
                    lastSellOrder.EntryTime > currentTrade.EntryTime AndAlso lastSellOrder.EntryPrice > currentTrade.EntryPrice Then
                    triggerPrice = currentTrade.EntryPrice
                    remark = "Breakeven"
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim lastBuyOrder As Trade = GetLastOrder(Trade.TradeExecutionDirection.Buy, currentMinuteCandlePayload)
                If lastBuyOrder IsNot Nothing AndAlso lastBuyOrder.TradeCurrentStatus <> Trade.TradeExecutionStatus.Close AndAlso
                    lastBuyOrder.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel AndAlso
                    lastBuyOrder.EntryTime > currentTrade.EntryTime AndAlso lastBuyOrder.EntryPrice < currentTrade.EntryPrice Then
                    triggerPrice = currentTrade.EntryPrice
                    remark = "Breakeven"
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim price As Decimal = Decimal.MinValue
            Dim remark As String = Nothing
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim lastSellOrder As Trade = GetLastOrder(Trade.TradeExecutionDirection.Sell, currentMinuteCandlePayload)
                If lastSellOrder IsNot Nothing AndAlso lastSellOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso
                    lastSellOrder.EntryTime < currentTrade.EntryTime Then
                    price = ConvertFloorCeling(lastSellOrder.PotentialStopLoss * 1.001, _parentStrategy.TickSize, RoundOfType.Celing)
                    remark = "Modified"
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim lastBuyOrder As Trade = GetLastOrder(Trade.TradeExecutionDirection.Buy, currentMinuteCandlePayload)
                If lastBuyOrder IsNot Nothing AndAlso lastBuyOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso
                    lastBuyOrder.EntryTime < currentTrade.EntryTime Then
                    price = ConvertFloorCeling(lastBuyOrder.PotentialStopLoss * 0.999, _parentStrategy.TickSize, RoundOfType.Floor)
                    remark = "Modified"
                End If
            End If
            If price <> Decimal.MinValue AndAlso price <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, price, remark)
            End If
        End If
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
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If direction = Trade.TradeExecutionDirection.Buy Then
                Dim buyPrice As Decimal = GetSlabBasedLevel(candle.PreviousCandlePayload.High, Trade.TradeExecutionDirection.Buy)
                ret = New Tuple(Of Boolean, Decimal, Payload, String)(True, buyPrice, candle.PreviousCandlePayload, "")
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                Dim sellPrice As Decimal = GetSlabBasedLevel(candle.PreviousCandlePayload.Low, Trade.TradeExecutionDirection.Sell)
                ret = New Tuple(Of Boolean, Decimal, Payload, String)(True, sellPrice, candle.PreviousCandlePayload, "")
            End If
        End If
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
End Class
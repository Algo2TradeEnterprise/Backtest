Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class LowStoplossSlabBasedStrategyRule
    Inherits StrategyRule

    Private _hkPayloads As Dictionary(Of Date, Payload)

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
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim quantity As Decimal = Me.LotSize
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim slPoint As Decimal = _slab
                Dim target As Decimal = _slab + 2 * buffer
                Dim remark As String = "Satellite"
                Dim anchorTrade As Trade = GetAnchorTrade(signal.Item4, currentTick)
                If anchorTrade Is Nothing OrElse (anchorTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close OrElse
                    anchorTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Cancel) Then
                    target = 100000000
                    remark = "Anchor"
                End If
                Dim lastOrder As Trade = GetLastOrder(signal.Item4, currentTick)
                If lastOrder Is Nothing OrElse (lastOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Close OrElse
                    lastOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Cancel) OrElse
                    (lastOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso lastOrder.Supporting2 = "Anchor" AndAlso
                    lastOrder.SignalCandle.PayloadDate <> signal.Item3.PayloadDate) Then
                    If remark = "Anchor" Then
                        If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                            If currentTick.Open < signal.Item2 + buffer Then
                                parameter = New PlaceOrderParameters With {
                                        .EntryPrice = signal.Item2 + buffer,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = quantity,
                                        .Stoploss = .EntryPrice - slPoint - 2 * buffer,
                                        .Target = .EntryPrice + target,
                                        .Buffer = buffer,
                                        .SignalCandle = signal.Item3,
                                        .OrderType = Trade.TypeOfOrder.SL,
                                        .Supporting1 = signal.Item3.PayloadDate.ToString("HH:mm:ss"),
                                        .Supporting2 = remark
                                    }
                            End If
                        ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                            If currentTick.Open > signal.Item2 - buffer Then
                                parameter = New PlaceOrderParameters With {
                                        .EntryPrice = signal.Item2 - buffer,
                                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                        .Quantity = quantity,
                                        .Stoploss = .EntryPrice + slPoint + 2 * buffer,
                                        .Target = .EntryPrice - target,
                                        .Buffer = buffer,
                                        .SignalCandle = signal.Item3,
                                        .OrderType = Trade.TypeOfOrder.SL,
                                        .Supporting1 = signal.Item3.PayloadDate.ToString("HH:mm:ss"),
                                        .Supporting2 = remark
                                    }
                            End If
                        End If
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
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryDirection = signal.Item4 AndAlso currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
            If ret Is Nothing Then
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    Dim buyLevel As Decimal = currentTrade.EntryPrice - currentTrade.EntryBuffer
                    If currentMinuteCandlePayload.PreviousCandlePayload.Low <= buyLevel - 2 * _slab Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    Dim sellLevel As Decimal = currentTrade.EntryPrice + currentTrade.EntryBuffer
                    If currentMinuteCandlePayload.PreviousCandlePayload.High >= sellLevel + 2 * _slab Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
            End If
            If ret Is Nothing Then
                Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
                If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                    Dim entryCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedTrade.EntryTime, _signalPayload))
                    If entryCandle.PayloadDate = currentMinuteCandlePayload.PayloadDate Then
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

    Private Function GetSlabBasedLevel(ByVal price As Decimal, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If direction = Trade.TradeExecutionDirection.Buy Then
            ret = Math.Ceiling(price / _slab) * _slab
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            ret = Math.Floor(price / _slab) * _slab
        End If
        Return ret
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayloads(candle.PayloadDate)
            If hkCandle.PreviousCandlePayload IsNot Nothing Then
                If hkCandle.PreviousCandlePayload.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                    Dim buyLevel As Decimal = GetSlabBasedLevel(hkCandle.PreviousCandlePayload.High, Trade.TradeExecutionDirection.Buy)
                    If hkCandle.PreviousCandlePayload.Low > buyLevel - 2 * _slab Then
                        ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, buyLevel, hkCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy)
                    End If
                ElseIf hkCandle.PreviousCandlePayload.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                    Dim sellLevel As Decimal = GetSlabBasedLevel(hkCandle.PreviousCandlePayload.Low, Trade.TradeExecutionDirection.Sell)
                    If hkCandle.PreviousCandlePayload.High < sellLevel + 2 * _slab Then
                        ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, sellLevel, hkCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell)
                    End If
                End If
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

    Private Function GetAnchorTrade(ByVal direction As Trade.TradeExecutionDirection, ByVal candle As Payload) As Trade
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
                                                                                  Return x.EntryDirection = direction AndAlso
                                                                                  x.Supporting2 = "Anchor"
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
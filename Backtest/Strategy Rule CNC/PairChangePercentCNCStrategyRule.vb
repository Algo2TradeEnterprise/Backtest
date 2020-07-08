Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class PairChangePercentCNCStrategyRule
    Inherits StrategyRule

    Public DummyCandle As Payload = Nothing
    Public Direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Public ChangePercentagePayloads As Dictionary(Of Date, Payload) = Nothing

    Private _diffPayloads As Dictionary(Of Date, Payload) = Nothing
    Private _firstSMAPayloads As Dictionary(Of Date, Decimal) = Nothing
    Private _firstBollingerHighPayloads As Dictionary(Of Date, Decimal) = Nothing
    Private _firstBollingerLowPayloads As Dictionary(Of Date, Decimal) = Nothing

    Private ReadOnly _controller As Boolean
    Private ReadOnly _stockType As Trade.TypeOfStock
    Private ReadOnly _openPrice As Decimal

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal canceller As CancellationTokenSource,
                   ByVal controller As Integer,
                   ByVal stockType As Trade.TypeOfStock,
                   ByVal openPrice As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, canceller)

        If controller = 1 Then
            _controller = True
        End If
        _stockType = stockType
        _openPrice = openPrice
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            Dim previousChangePayload As Payload = Nothing
            For Each runningPayload In _signalPayload
                Dim close As Decimal = _openPrice
                If close <> Decimal.MinValue Then
                    Dim change As Payload = New Payload(Payload.CandleDataSource.Calculated) With
                        {
                         .PayloadDate = runningPayload.Key,
                         .Open = Math.Round(((runningPayload.Value.Close / close) - 1) * 100, 3),
                         .High = Math.Round(((runningPayload.Value.Close / close) - 1) * 100, 3),
                         .Low = Math.Round(((runningPayload.Value.Close / close) - 1) * 100, 3),
                         .Close = Math.Round(((runningPayload.Value.Close / close) - 1) * 100, 3),
                         .Volume = 1,
                         .PreviousCandlePayload = previousChangePayload
                        }

                    If ChangePercentagePayloads Is Nothing Then ChangePercentagePayloads = New Dictionary(Of Date, Payload)
                    ChangePercentagePayloads.Add(runningPayload.Key, change)
                    previousChangePayload = change
                End If
            Next
        End If
    End Sub

    Public Overrides Sub CompletePairProcessing()
        MyBase.CompletePairProcessing()

        Dim myPair As PairChangePercentCNCStrategyRule = Me.AnotherPairInstrument
        If Me.ChangePercentagePayloads IsNot Nothing AndAlso Me.ChangePercentagePayloads.Count > 0 AndAlso
            myPair.ChangePercentagePayloads IsNot Nothing AndAlso myPair.ChangePercentagePayloads.Count > 0 Then
            Dim previousDiffPayload As Payload = Nothing
            For Each runningPayload In Me.ChangePercentagePayloads.Keys
                If myPair.ChangePercentagePayloads.ContainsKey(runningPayload) Then
                    Dim diff As Payload = New Payload(Payload.CandleDataSource.Calculated) With
                    {
                     .PayloadDate = runningPayload,
                     .Open = Me.ChangePercentagePayloads(runningPayload).Close - myPair.ChangePercentagePayloads(runningPayload).Close,
                     .High = Me.ChangePercentagePayloads(runningPayload).Close - myPair.ChangePercentagePayloads(runningPayload).Close,
                     .Low = Me.ChangePercentagePayloads(runningPayload).Close - myPair.ChangePercentagePayloads(runningPayload).Close,
                     .Close = Me.ChangePercentagePayloads(runningPayload).Close - myPair.ChangePercentagePayloads(runningPayload).Close,
                     .Volume = 1,
                     .PreviousCandlePayload = previousDiffPayload
                    }

                    If _diffPayloads Is Nothing Then _diffPayloads = New Dictionary(Of Date, Payload)
                    _diffPayloads.Add(runningPayload, diff)
                End If
            Next

            If _diffPayloads IsNot Nothing AndAlso _diffPayloads.Count > 0 Then
                Indicator.BollingerBands.CalculateBollingerBands(50, Payload.PayloadFields.Close, 4, _diffPayloads, _firstBollingerHighPayloads, _firstBollingerLowPayloads, _firstSMAPayloads)
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Me.DummyCandle = currentTick
        If Me.ForceTakeTrade AndAlso Not _controller Then
            Dim quantity As Integer = 1
            If _stockType = Trade.TypeOfStock.Cash Then
                If currentTick.Open > CType(Me.AnotherPairInstrument, PairChangePercentCNCStrategyRule).DummyCandle.Open Then
                    quantity = Me.LotSize
                Else
                    Dim multiplier As Decimal = CType(Me.AnotherPairInstrument, PairChangePercentCNCStrategyRule).DummyCandle.Open / currentTick.Open
                    quantity = Math.Floor(multiplier * CType(Me.AnotherPairInstrument, PairChangePercentCNCStrategyRule).LotSize)
                End If
            Else
                quantity = Me.LotSize
            End If

            Dim parameter As PlaceOrderParameters = Nothing
            If Me.Direction = Trade.TradeExecutionDirection.Buy Then
                parameter = New PlaceOrderParameters With {
                                    .EntryPrice = currentTick.Open,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - 100000,
                                    .Target = .EntryPrice + 100000,
                                    .Buffer = 0,
                                    .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.Market
                                }
            ElseIf Me.Direction = Trade.TradeExecutionDirection.Sell Then
                parameter = New PlaceOrderParameters With {
                                    .EntryPrice = currentTick.Open,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + 100000,
                                    .Target = .EntryPrice - 100000,
                                    .Buffer = 0,
                                    .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.Market
                                }
            End If

            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            Me.ForceTakeTrade = False
        ElseIf _controller Then
            Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
            Dim parameter As PlaceOrderParameters = Nothing
            If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
                Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.CNC) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.CNC) AndAlso
                currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
                Dim signalCandle As Payload = Nothing
                Dim signal As Tuple(Of Boolean, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                End If
                If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                    Dim quantity As Integer = 1
                    If _stockType = Trade.TypeOfStock.Cash Then
                        If currentTick.Open > CType(Me.AnotherPairInstrument, PairChangePercentCNCStrategyRule).DummyCandle.Open Then
                            quantity = Me.LotSize
                        Else
                            Dim multiplier As Decimal = CType(Me.AnotherPairInstrument, PairChangePercentCNCStrategyRule).DummyCandle.Open / currentTick.Open
                            quantity = Math.Floor(multiplier * CType(Me.AnotherPairInstrument, PairChangePercentCNCStrategyRule).LotSize)
                        End If
                    Else
                        quantity = Me.LotSize
                    End If
                    Dim entryDiff As Decimal = _diffPayloads(signalCandle.PayloadDate).Close
                    Dim entryHighBollinger As Decimal = _firstBollingerHighPayloads(signalCandle.PayloadDate)
                    Dim entryLowBollinger As Decimal = _firstBollingerLowPayloads(signalCandle.PayloadDate)
                    Dim remark As String = ""
                    If entryDiff > entryHighBollinger Then
                        remark = "+ SD"
                    ElseIf entryDiff < entryLowBollinger Then
                        remark = "- SD"
                    End If

                    If signal.Item2 = Trade.TradeExecutionDirection.Buy Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = currentTick.Open,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - 100000,
                                    .Target = .EntryPrice + 100000,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = entryDiff,
                                    .Supporting2 = entryHighBollinger,
                                    .Supporting3 = entryLowBollinger,
                                    .Supporting4 = remark
                                }
                        CType(Me.AnotherPairInstrument, PairChangePercentCNCStrategyRule).Direction = Trade.TradeExecutionDirection.Sell
                        Me.AnotherPairInstrument.ForceTakeTrade = True
                    ElseIf signal.Item2 = Trade.TradeExecutionDirection.Sell Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = currentTick.Open,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + 100000,
                                    .Target = .EntryPrice - 100000,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = entryDiff,
                                    .Supporting2 = entryHighBollinger,
                                    .Supporting3 = entryLowBollinger,
                                    .Supporting4 = remark
                                }
                        CType(Me.AnotherPairInstrument, PairChangePercentCNCStrategyRule).Direction = Trade.TradeExecutionDirection.Buy
                        Me.AnotherPairInstrument.ForceTakeTrade = True
                    End If
                End If
            End If
            If parameter IsNot Nothing Then
                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Me.DummyCandle = currentTick
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If Me.ForceCancelTrade AndAlso Not _controller Then
                ret = New Tuple(Of Boolean, String)(True, "Normal Force Exit")
                Me.ForceCancelTrade = False
            ElseIf _controller Then
                Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                Dim entryDiff As Decimal = currentTrade.Supporting1
                Dim entryHighBollinger As Decimal = currentTrade.Supporting2
                Dim entryLowBollinger As Decimal = currentTrade.Supporting3
                Dim currentDiff As Decimal = _diffPayloads(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate).Close
                Dim currentHighBollinger As Decimal = _firstBollingerHighPayloads(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                Dim currentLowBollinger As Decimal = _firstBollingerLowPayloads(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                If entryDiff > entryHighBollinger Then
                    If currentDiff < currentLowBollinger Then
                        ret = New Tuple(Of Boolean, String)(True, String.Format("Normal Exit at -SD"))
                        Me.AnotherPairInstrument.ForceCancelTrade = True
                    End If
                ElseIf entryDiff < entryLowBollinger Then
                    If currentDiff > currentHighBollinger Then
                        ret = New Tuple(Of Boolean, String)(True, String.Format("Normal Exit at +SD"))
                        Me.AnotherPairInstrument.ForceCancelTrade = True
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection) = Nothing
        Dim eligibleToTakeTrade As Boolean = False
        If CType(Me.AnotherPairInstrument, PairChangePercentCNCStrategyRule).DummyCandle IsNot Nothing Then
            Dim pairlastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(CType(Me.AnotherPairInstrument, PairChangePercentCNCStrategyRule).DummyCandle, Trade.TypeOfTrade.CNC)
            If pairlastExecutedOrder IsNot Nothing Then
                If pairlastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                    eligibleToTakeTrade = True
                End If
            Else
                eligibleToTakeTrade = True
            End If
        End If
        If eligibleToTakeTrade Then
            If _diffPayloads.ContainsKey(candle.PreviousCandlePayload.PayloadDate) Then
                Dim diff As Decimal = _diffPayloads(candle.PreviousCandlePayload.PayloadDate).Close
                Dim highBollinger As Decimal = _firstBollingerHighPayloads(candle.PreviousCandlePayload.PayloadDate)
                Dim lowBollinger As Decimal = _firstBollingerLowPayloads(candle.PreviousCandlePayload.PayloadDate)
                If diff > highBollinger Then
                    ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Sell)
                ElseIf diff < lowBollinger Then
                    ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Buy)
                End If
            End If
        End If
        Return ret
    End Function
End Class
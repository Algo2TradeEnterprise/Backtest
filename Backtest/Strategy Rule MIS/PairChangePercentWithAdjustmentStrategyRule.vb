Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PairChangePercentWithAdjustmentStrategyRule
    Inherits StrategyRule

    Public DummyCandle As Payload = Nothing
    Public Direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Public ChangePercentagePayloads As Dictionary(Of Date, Payload) = Nothing

    Private _diffPayloads As Dictionary(Of Date, Payload) = Nothing
    Private _firstSMAPayloads As Dictionary(Of Date, Decimal) = Nothing
    Private _firstBollingerHighPayloads As Dictionary(Of Date, Decimal) = Nothing
    Private _firstBollingerLowPayloads As Dictionary(Of Date, Decimal) = Nothing
    Private _atrPayloads As Dictionary(Of Date, Decimal) = Nothing
    Private _emaPayloads As Dictionary(Of Date, Decimal) = Nothing
    Private _swingPayloads As Dictionary(Of Date, Indicator.Swing) = Nothing
    Private _exitDone As Boolean = False
    Private ReadOnly _buffer As Decimal = 2.5

    Private ReadOnly _controller As Boolean
    Private ReadOnly _stockType As Trade.TypeOfStock

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal canceller As CancellationTokenSource,
                   ByVal controller As Integer,
                   ByVal stockType As Trade.TypeOfStock)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, canceller)

        If controller = 1 Then
            _controller = True
        End If
        _stockType = stockType
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayloads)
            Indicator.EMA.CalculateEMA(13, Payload.PayloadFields.Close, _signalPayload, _emaPayloads)
            Indicator.SwingHighLow.CalculateSwingHighLow(_signalPayload, True, _swingPayloads)

            Dim firstCandleOfTheDay As Payload = _signalPayload.Where(Function(x)
                                                                          Return x.Key.Date = _tradingDate.Date
                                                                      End Function).OrderBy(Function(y)
                                                                                                Return y.Key
                                                                                            End Function).FirstOrDefault.Value
            Dim previousTradingDay As Date = firstCandleOfTheDay.PreviousCandlePayload.PayloadDate
            Dim previousDayLastCandleClose As Decimal = firstCandleOfTheDay.PreviousCandlePayload.Close
            Dim firstCandleOfThePreviousDay As Payload = _signalPayload.Where(Function(x)
                                                                                  Return x.Key.Date = previousTradingDay.Date
                                                                              End Function).OrderBy(Function(y)
                                                                                                        Return y.Key
                                                                                                    End Function).FirstOrDefault.Value
            Dim prePreviousDayLastCandleClose As Decimal = firstCandleOfThePreviousDay.PreviousCandlePayload.Close

            Dim previousChangePayload As Payload = Nothing
            For Each runningPayload In _signalPayload
                If runningPayload.Key >= previousTradingDay.Date Then
                    'Dim close As Decimal = Decimal.MinValue
                    'If runningPayload.Key.Date = previousTradingDay.Date Then
                    '    close = prePreviousDayLastCandleClose
                    'ElseIf runningPayload.Key.Date = _tradingDate.Date Then
                    '    close = previousDayLastCandleClose
                    'End If
                    Dim close As Decimal = firstCandleOfThePreviousDay.Close
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
                End If
            Next
        End If
    End Sub

    Public Overrides Sub CompletePairProcessing()
        MyBase.CompletePairProcessing()

        If _controller Then
            Dim myPair As PairChangePercentWithAdjustmentStrategyRule = Me.AnotherPairInstrument
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
                    Indicator.BollingerBands.CalculateBollingerBands(50, Payload.PayloadFields.Close, 3, _diffPayloads, _firstBollingerHighPayloads, _firstBollingerLowPayloads, _firstSMAPayloads)

                    For Each runningPayload In _diffPayloads
                        If runningPayload.Key.Date = _tradingDate.Date Then
                            Dim diff As Decimal = runningPayload.Value.Close
                            Dim plusSD As Decimal = _firstBollingerHighPayloads(runningPayload.Key)
                            Dim minusSD As Decimal = _firstBollingerLowPayloads(runningPayload.Key)
                            Dim sma As Decimal = _firstSMAPayloads(runningPayload.Key)

                            Console.WriteLine(String.Format("{0},{1},{2},{3},{4},{5},{6}",
                                                            runningPayload.Key.ToString("HH:mm:ss"),
                                                            Me.ChangePercentagePayloads(runningPayload.Key).Close,
                                                            myPair.ChangePercentagePayloads(runningPayload.Key).Close,
                                                            diff, sma, plusSD, minusSD))
                        End If
                    Next
                    Console.WriteLine("--------------------------------------")
                End If
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Me.DummyCandle = currentTick
        Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
        If lastExecutedOrder IsNot Nothing AndAlso Not _exitDone AndAlso Not _parentStrategy.IsTradeOpen(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item3
            End If
            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim quantity As Integer = 1
                If _stockType = Trade.TypeOfStock.Cash Then
                    If currentTick.Open > CType(Me.AnotherPairInstrument, PairChangePercentWithAdjustmentStrategyRule).DummyCandle.Open Then
                        quantity = Me.LotSize
                    Else
                        Dim multiplier As Decimal = CType(Me.AnotherPairInstrument, PairChangePercentWithAdjustmentStrategyRule).DummyCandle.Open / currentTick.Open
                        quantity = Math.Floor(multiplier * CType(Me.AnotherPairInstrument, PairChangePercentWithAdjustmentStrategyRule).LotSize)
                    End If
                Else
                    quantity = Me.LotSize
                End If

                Dim parameter As PlaceOrderParameters = Nothing
                If signal.Item2 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    If _tradingSymbol.Contains("BANKNIFTY") Then buffer = _buffer

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalCandle.High + buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - 100000,
                                .Target = .EntryPrice + 100000,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = lastExecutedOrder.Supporting2,
                                .Supporting3 = lastExecutedOrder.Supporting3,
                                .Supporting4 = lastExecutedOrder.Supporting4
                            }
                ElseIf signal.Item2 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    If _tradingSymbol.Contains("BANKNIFTY") Then buffer = _buffer

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalCandle.Low - buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + 100000,
                                .Target = .EntryPrice - 100000,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = lastExecutedOrder.Supporting2,
                                .Supporting3 = lastExecutedOrder.Supporting3,
                                .Supporting4 = lastExecutedOrder.Supporting4
                            }
                End If
                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            End If
        Else
            If Me.ForceTakeTrade AndAlso Not _controller Then
                Dim quantity As Integer = 1
                If _stockType = Trade.TypeOfStock.Cash Then
                    If currentTick.Open > CType(Me.AnotherPairInstrument, PairChangePercentWithAdjustmentStrategyRule).DummyCandle.Open Then
                        quantity = Me.LotSize
                    Else
                        Dim multiplier As Decimal = CType(Me.AnotherPairInstrument, PairChangePercentWithAdjustmentStrategyRule).DummyCandle.Open / currentTick.Open
                        quantity = Math.Floor(multiplier * CType(Me.AnotherPairInstrument, PairChangePercentWithAdjustmentStrategyRule).LotSize)
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
                    Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
                    currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso Not _exitDone Then
                    Dim signalCandle As Payload = Nothing
                    Dim signal As Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
                    If signal IsNot Nothing AndAlso signal.Item1 Then
                        signalCandle = signal.Item3
                    End If
                    If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                        Dim quantity As Integer = 1
                        If _stockType = Trade.TypeOfStock.Cash Then
                            If currentTick.Open > CType(Me.AnotherPairInstrument, PairChangePercentWithAdjustmentStrategyRule).DummyCandle.Open Then
                                quantity = Me.LotSize
                            Else
                                Dim multiplier As Decimal = CType(Me.AnotherPairInstrument, PairChangePercentWithAdjustmentStrategyRule).DummyCandle.Open / currentTick.Open
                                quantity = Math.Floor(multiplier * CType(Me.AnotherPairInstrument, PairChangePercentWithAdjustmentStrategyRule).LotSize)
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
                                        .Supporting1 = remark,
                                        .Supporting2 = entryDiff,
                                        .Supporting3 = entryHighBollinger,
                                        .Supporting4 = entryLowBollinger
                                    }
                            CType(Me.AnotherPairInstrument, PairChangePercentWithAdjustmentStrategyRule).Direction = Trade.TradeExecutionDirection.Sell
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
                                        .Supporting1 = remark,
                                        .Supporting2 = entryDiff,
                                        .Supporting3 = entryHighBollinger,
                                        .Supporting4 = entryLowBollinger
                                    }
                            CType(Me.AnotherPairInstrument, PairChangePercentWithAdjustmentStrategyRule).Direction = Trade.TradeExecutionDirection.Buy
                            Me.AnotherPairInstrument.ForceTakeTrade = True
                        End If
                    End If
                End If
                If parameter IsNot Nothing Then
                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                End If
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
                'Me.ForceCancelTrade = False
            ElseIf _controller Then
                Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                Dim entryDiff As Decimal = currentTrade.Supporting2
                Dim entryHighBollinger As Decimal = currentTrade.Supporting3
                Dim entryLowBollinger As Decimal = currentTrade.Supporting4
                Dim currentDiff As Decimal = _diffPayloads(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate).Close
                Dim currentHighBollinger As Decimal = _firstBollingerHighPayloads(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                Dim currentLowBollinger As Decimal = _firstBollingerLowPayloads(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                If entryDiff > entryHighBollinger Then
                    If currentDiff < currentLowBollinger Then
                        ret = New Tuple(Of Boolean, String)(True, String.Format("Normal Exit at -SD"))
                        Me.AnotherPairInstrument.ForceCancelTrade = True
                        _exitDone = True
                    End If
                ElseIf entryDiff < entryLowBollinger Then
                    If currentDiff > currentHighBollinger Then
                        ret = New Tuple(Of Boolean, String)(True, String.Format("Normal Exit at +SD"))
                        Me.AnotherPairInstrument.ForceCancelTrade = True
                        _exitDone = True
                    End If
                End If
            End If
        ElseIf currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            If _exitDone Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid signal")
            Else
                Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                Dim signal As Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
                If signal IsNot Nothing AndAlso signal.Item3.PayloadDate <> currentTrade.SignalCandle.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid signal")
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload) = Nothing
        Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
        If lastExecutedOrder Is Nothing Then
            If _diffPayloads.ContainsKey(candle.PreviousCandlePayload.PayloadDate) Then
                Dim diff As Decimal = _diffPayloads(candle.PreviousCandlePayload.PayloadDate).Close
                Dim highBollinger As Decimal = _firstBollingerHighPayloads(candle.PreviousCandlePayload.PayloadDate)
                Dim lowBollinger As Decimal = _firstBollingerLowPayloads(candle.PreviousCandlePayload.PayloadDate)
                If diff > highBollinger OrElse diff < lowBollinger Then
                    Dim myChange As Decimal = Me.ChangePercentagePayloads(candle.PreviousCandlePayload.PayloadDate).Close
                    Dim myPairChange As Decimal = CType(Me.AnotherPairInstrument, PairChangePercentWithAdjustmentStrategyRule).ChangePercentagePayloads(candle.PreviousCandlePayload.PayloadDate).Close
                    If myChange > myPairChange Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload)(True, Trade.TradeExecutionDirection.Sell, candle.PreviousCandlePayload)
                    ElseIf myChange < myPairChange Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload)(True, Trade.TradeExecutionDirection.Buy, candle.PreviousCandlePayload)
                    End If
                End If
            End If
        ElseIf lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If lastExecutedOrder.SignalCandle.PayloadDate < candle.PreviousCandlePayload.PayloadDate Then
                If lastExecutedOrder.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    Dim swing As Indicator.Swing = _swingPayloads(candle.PreviousCandlePayload.PayloadDate)
                    If swing.SwingHighTime > lastExecutedOrder.EntryTime Then
                        Dim ema As Decimal = _emaPayloads(swing.SwingHighTime)
                        If swing.SwingHigh > ema Then
                            Dim signalCandle As Payload = _signalPayload(swing.SwingHighTime)
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                            If _tradingSymbol.Contains("BANKNIFTY") Then buffer = _buffer
                            Dim atr As Decimal = GetHighestATRofTheDay(candle.PreviousCandlePayload)
                            If lastExecutedOrder.EntryPrice - (signalCandle.High + buffer) >= atr Then
                                ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload)(True, Trade.TradeExecutionDirection.Buy, signalCandle)
                            End If
                        End If
                    End If
                ElseIf lastExecutedOrder.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    Dim swing As Indicator.Swing = _swingPayloads(candle.PreviousCandlePayload.PayloadDate)
                    If swing.SwingLowTime > lastExecutedOrder.EntryTime Then
                        Dim ema As Decimal = _emaPayloads(swing.SwingHighTime)
                        If swing.SwingLow < ema Then
                            Dim signalCandle As Payload = _signalPayload(swing.SwingLowTime)
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                            If _tradingSymbol.Contains("BANKNIFTY") Then buffer = _buffer
                            Dim atr As Decimal = GetHighestATRofTheDay(candle.PreviousCandlePayload)
                            If (signalCandle.Low - buffer) - lastExecutedOrder.EntryPrice >= atr Then
                                ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload)(True, Trade.TradeExecutionDirection.Sell, signalCandle)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestATRofTheDay(ByVal signalCandle As Payload)
        Dim ret As Decimal = Decimal.MinValue
        ret = _atrPayloads.Max(Function(x)
                                   If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                       Return x.Value
                                   Else
                                       Return Decimal.MinValue
                                   End If
                               End Function)
        Return ret
    End Function
End Class
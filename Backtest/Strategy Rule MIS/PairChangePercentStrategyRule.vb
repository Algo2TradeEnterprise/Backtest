Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PairChangePercentStrategyRule
    Inherits StrategyRule

    Public DummyCandle As Payload = Nothing
    Public Direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Public ChangePercentagePayloads As Dictionary(Of Date, Payload) = Nothing

    Private _diffPayloads As Dictionary(Of Date, Payload) = Nothing
    Private _smaPayloads As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerHighPayloads As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerLowPayloads As Dictionary(Of Date, Decimal) = Nothing

    Private ReadOnly _controller As Boolean

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal canceller As CancellationTokenSource,
                   ByVal controller As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, canceller)

        If controller = 1 Then
            _controller = True
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
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
                    Dim close As Decimal = Decimal.MinValue
                    If runningPayload.Key.Date = previousTradingDay.Date Then
                        close = prePreviousDayLastCandleClose
                    ElseIf runningPayload.Key.Date = _tradingDate.Date Then
                        close = previousDayLastCandleClose
                    End If
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

        Dim myPair As PairChangePercentStrategyRule = Me.AnotherPairInstrument
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
                Indicator.BollingerBands.CalculateBollingerBands(50, Payload.PayloadFields.Close, 2, _diffPayloads, _bollingerHighPayloads, _bollingerLowPayloads, _smaPayloads)
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
            If currentTick.Open > CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).DummyCandle.Open Then
                quantity = Me.LotSize
            Else
                Dim multiplier As Decimal = CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).DummyCandle.Open / currentTick.Open
                quantity = Math.Floor(multiplier * CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).LotSize)
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
                currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

                Dim signalCandle As Payload = Nothing
                Dim signal As Tuple(Of Boolean, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                End If
                If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                    Dim quantity As Integer = 1
                    If currentTick.Open > CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).DummyCandle.Open Then
                        quantity = Me.LotSize
                    Else
                        Dim multiplier As Decimal = CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).DummyCandle.Open / currentTick.Open
                        quantity = Math.Floor(multiplier * CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).LotSize)
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
                                        .OrderType = Trade.TypeOfOrder.Market
                                    }
                        CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).Direction = Trade.TradeExecutionDirection.Sell
                    ElseIf signal.Item2 = Trade.TradeExecutionDirection.Sell Then
                        parameter = New PlaceOrderParameters With {
                                        .EntryPrice = currentTick.Open,
                                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                        .Quantity = quantity,
                                        .Stoploss = .EntryPrice + 100000,
                                        .Target = .EntryPrice - 100000,
                                        .Buffer = 0,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.Market
                                    }
                        CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).Direction = Trade.TradeExecutionDirection.Buy
                    End If
                End If
            End If
            If parameter IsNot Nothing Then
                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                Me.AnotherPairInstrument.ForceTakeTrade = True
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
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
        If CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).DummyCandle IsNot Nothing Then
            Dim pairlastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).DummyCandle, Trade.TypeOfTrade.MIS)
            If pairlastExecutedOrder IsNot Nothing Then
                If pairlastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                    eligibleToTakeTrade = True
                End If
            Else
                eligibleToTakeTrade = True
            End If
        End If
        If eligibleToTakeTrade Then
            Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
            If lastExecutedOrder Is Nothing OrElse (lastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Close AndAlso candle.PayloadDate > lastExecutedOrder.ExitTime) Then
                If _diffPayloads.ContainsKey(candle.PreviousCandlePayload.PayloadDate) Then
                    Dim diff As Decimal = _diffPayloads(candle.PreviousCandlePayload.PayloadDate).Close
                    Dim highBollinger As Decimal = _bollingerHighPayloads(candle.PreviousCandlePayload.PayloadDate)
                    Dim lowBollinger As Decimal = _bollingerLowPayloads(candle.PreviousCandlePayload.PayloadDate)
                    If diff > highBollinger OrElse diff < lowBollinger Then
                        Dim myChange As Decimal = Me.ChangePercentagePayloads(candle.PreviousCandlePayload.PayloadDate).Close
                        Dim myPairChange As Decimal = CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).ChangePercentagePayloads(candle.PreviousCandlePayload.PayloadDate).Close
                        If myChange > myPairChange Then
                            ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Sell)
                        ElseIf myChange < myPairChange Then
                            ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Buy)
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class

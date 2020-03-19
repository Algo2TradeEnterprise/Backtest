Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PairBollingerDifferenceStrategyRule
    Inherits StrategyRule

    Private ReadOnly _minDifferenceForEntry As Decimal = 2
    Private ReadOnly _maxDifferenceForExit As Decimal = 0.5


    Public Direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Public ChangePercentagePayloads As Dictionary(Of Date, Payload) = Nothing
    Public LastTick As Payload = Nothing

    Private _ChangePercentageDifferencePayloads As Dictionary(Of Date, Payload) = Nothing
    Private _bollingerHighPayloads As Dictionary(Of Date, Decimal)
    Private _bollingerLowPayloads As Dictionary(Of Date, Decimal)
    Private _smaPayloads As Dictionary(Of Date, Decimal)

    Private _previousTradingDay As Date

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
            _previousTradingDay = _parentStrategy.Cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Cash, _tradingDate)

            Dim previousDayFirstCandleClose As Decimal = _signalPayload.Where(Function(x)
                                                                                  Return x.Key.Date = _previousTradingDay.Date
                                                                              End Function).OrderBy(Function(y)
                                                                                                        Return y.Key
                                                                                                    End Function).FirstOrDefault.Value.Close

            Dim previousChangePayload As Payload = Nothing
            For Each runningPayload In _signalPayload
                If ChangePercentagePayloads Is Nothing Then ChangePercentagePayloads = New Dictionary(Of Date, Payload)
                Dim change As Payload = New Payload(Payload.CandleDataSource.Calculated) With
                    {
                     .PayloadDate = runningPayload.Key,
                     .Open = Math.Round(((runningPayload.Value.Close / previousDayFirstCandleClose) - 1) * 100, 3),
                     .High = Math.Round(((runningPayload.Value.Close / previousDayFirstCandleClose) - 1) * 100, 3),
                     .Low = Math.Round(((runningPayload.Value.Close / previousDayFirstCandleClose) - 1) * 100, 3),
                     .Close = Math.Round(((runningPayload.Value.Close / previousDayFirstCandleClose) - 1) * 100, 3),
                     .Volume = 1,
                     .PreviousCandlePayload = previousChangePayload
                    }
                ChangePercentagePayloads.Add(runningPayload.Key, change)
                previousChangePayload = change
            Next
        End If
    End Sub

    Public Overrides Sub CompletePairProcessing()
        MyBase.CompletePairProcessing()
        If _controller Then
            Dim myPair As PairBollingerDifferenceStrategyRule = Me.AnotherPairInstrument
            Dim previousChangePayload As Payload = Nothing
            For Each runningPayload In _signalPayload
                If _ChangePercentageDifferencePayloads Is Nothing Then _ChangePercentageDifferencePayloads = New Dictionary(Of Date, Payload)
                Dim change As Payload = New Payload(Payload.CandleDataSource.Calculated) With
                {
                 .PayloadDate = runningPayload.Key,
                 .Open = Math.Round(Math.Abs(Me.ChangePercentagePayloads(runningPayload.Key).Close - myPair.ChangePercentagePayloads(runningPayload.Key).Close), 3),
                 .Low = Math.Round(Math.Abs(Me.ChangePercentagePayloads(runningPayload.Key).Close - myPair.ChangePercentagePayloads(runningPayload.Key).Close), 3),
                 .High = Math.Round(Math.Abs(Me.ChangePercentagePayloads(runningPayload.Key).Close - myPair.ChangePercentagePayloads(runningPayload.Key).Close), 3),
                 .Close = Math.Round(Math.Abs(Me.ChangePercentagePayloads(runningPayload.Key).Close - myPair.ChangePercentagePayloads(runningPayload.Key).Close), 3),
                 .Volume = 1,
                 .PreviousCandlePayload = previousChangePayload
                }
                _ChangePercentageDifferencePayloads.Add(runningPayload.Key, change)
                previousChangePayload = change
            Next

            If _ChangePercentageDifferencePayloads IsNot Nothing AndAlso _ChangePercentageDifferencePayloads.Count > 0 Then
                Dim requiredPayload As IEnumerable(Of KeyValuePair(Of Date, Payload)) = _ChangePercentageDifferencePayloads.Where(Function(x)
                                                                                                                                      Return x.Key.Date >= _previousTradingDay.Date
                                                                                                                                  End Function)
                Dim twoDaysPayload As Dictionary(Of Date, Payload) = Nothing
                If requiredPayload IsNot Nothing AndAlso requiredPayload.Count > 0 Then
                    For Each runningPayload In requiredPayload
                        If twoDaysPayload Is Nothing Then twoDaysPayload = New Dictionary(Of Date, Payload)
                        twoDaysPayload.Add(runningPayload.Key, runningPayload.Value)
                    Next
                End If
                If twoDaysPayload IsNot Nothing AndAlso twoDaysPayload.Count > 0 Then
                    Indicator.BollingerBands.CalculateBollingerBands(50, Payload.PayloadFields.Close, 3, twoDaysPayload,
                                                                 _bollingerHighPayloads, _bollingerLowPayloads, _smaPayloads)

                    For Each runningPayload In twoDaysPayload.Keys
                        If runningPayload.Date = _tradingDate.Date Then
                            Console.WriteLine(String.Format("{0},{1},{2},{3},{4},{5},{6}",
                                                            runningPayload.ToString("dd-MM-yyyy HH:mm:ss"),
                                                            Me.ChangePercentagePayloads(runningPayload).Close,
                                                            myPair.ChangePercentagePayloads(runningPayload).Close,
                                                            _ChangePercentageDifferencePayloads(runningPayload).Close,
                                                            _smaPayloads(runningPayload),
                                                            _bollingerHighPayloads(runningPayload),
                                                            _bollingerLowPayloads(runningPayload)))
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If Me.ForceTakeTrade AndAlso Not _controller Then
            Dim myPair As PairDifferenceStrategyRule = Me.AnotherPairInstrument
            Dim quantity As Integer = Math.Ceiling(myPair.LotSize * (myPair.LastTick.Open / Me.LastTick.Open))

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
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = "Force Entry"
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
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = "Force Entry"
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
                Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                End If
                If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                    Dim quantity As Integer = Me.LotSize

                    If signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy Then
                        parameter = New PlaceOrderParameters With {
                                        .EntryPrice = currentTick.Open,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = quantity,
                                        .Stoploss = .EntryPrice - 100000,
                                        .Target = .EntryPrice + 100000,
                                        .Buffer = 0,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting1 = "Normal Entry"
                                    }
                        CType(Me.AnotherPairInstrument, PairDifferenceStrategyRule).Direction = Trade.TradeExecutionDirection.Sell
                    ElseIf signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell Then
                        parameter = New PlaceOrderParameters With {
                                        .EntryPrice = currentTick.Open,
                                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                        .Quantity = quantity,
                                        .Stoploss = .EntryPrice + 100000,
                                        .Target = .EntryPrice - 100000,
                                        .Buffer = 0,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting1 = "Normal Entry"
                                    }
                        CType(Me.AnotherPairInstrument, PairDifferenceStrategyRule).Direction = Trade.TradeExecutionDirection.Buy
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
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If Me.ForceCancelTrade AndAlso Not _controller Then
                ret = New Tuple(Of Boolean, String)(True, "Force Exit")
                Me.ForceCancelTrade = False
            ElseIf _controller Then
                Dim myPair As PairDifferenceStrategyRule = Me.AnotherPairInstrument
                'If Math.Abs(myChange - myPairChange) <= _maxDifferenceForExit Then
                '    ret = New Tuple(Of Boolean, String)(True, String.Format("Normal Exit as Difference({0}) is less than {1}", Math.Round(Math.Abs(myChange - myPairChange), 2), _maxDifferenceForExit))
                '    Me.AnotherPairInstrument.ForceCancelTrade = True
                'End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection) = Nothing
        Dim myPair As PairDifferenceStrategyRule = Me.AnotherPairInstrument
        If myPair.LastTick IsNot Nothing Then
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(candle, Trade.TypeOfTrade.CNC)
            If lastExecutedTrade Is Nothing OrElse lastExecutedTrade.ExitTime < candle.PayloadDate Then
                'If Math.Abs(myChange - myPairChange) >= _minDifferenceForEntry Then
                '    If myChange < myPairChange Then
                '        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Buy)
                '    Else
                '        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Sell)
                '    End If
                'End If
            End If
        End If
        Return ret
    End Function
End Class

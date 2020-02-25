Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PairStrategyRule
    Inherits StrategyRule

    Public Direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Public ChangePercentagePayloads As Dictionary(Of Date, Payload) = Nothing

    Private _swingHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _swingLowPayload As Dictionary(Of Date, Decimal) = Nothing

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, canceller)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            Dim previousTradingDay As Date = _parentStrategy.Cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Cash, _tradingDate)

            Dim previousDayFirstCandleClose As Decimal = _signalPayload.Where(Function(x)
                                                                                  Return x.Key.Date = previousTradingDay.Date
                                                                              End Function).OrderBy(Function(y)
                                                                                                        Return y.Key
                                                                                                    End Function).FirstOrDefault.Value.Close

            Dim previousChangePayload As Payload = Nothing
            For Each runningPayload In _signalPayload
                If ChangePercentagePayloads Is Nothing Then ChangePercentagePayloads = New Dictionary(Of Date, Payload)
                Dim change As Payload = New Payload(Payload.CandleDataSource.Calculated) With
                    {
                     .PayloadDate = runningPayload.Key,
                     .High = Math.Round(((runningPayload.Value.Close / previousDayFirstCandleClose) - 1) * 100, 3),
                     .Low = Math.Round(((runningPayload.Value.Close / previousDayFirstCandleClose) - 1) * 100, 3),
                     .Volume = 1,
                     .PreviousCandlePayload = previousChangePayload
                    }
                ChangePercentagePayloads.Add(runningPayload.Key, change)
                previousChangePayload = change
            Next

            Indicator.SwingHighLow.CalculateSwingHighLow(ChangePercentagePayloads, True, _swingHighPayload, _swingLowPayload)
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        'If Me.ForceTakeTrade Then
        '    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromInvestment(Me.LotSize, 10000, currentTick.Open, _parentStrategy.StockType, True)
        '    Dim parameter As PlaceOrderParameters = Nothing
        '    'If _direction = Trade.TradeExecutionDirection.Buy Then
        '    '    parameter = New PlaceOrderParameters With {
        '    '                            .EntryPrice = currentTick.Open,
        '    '                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
        '    '                            .Quantity = quantity,
        '    '                            .Stoploss = .EntryPrice - 100000,
        '    '                            .Target = .EntryPrice + 100000,
        '    '                            .Buffer = 0,
        '    '                            .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
        '    '                            .OrderType = Trade.TypeOfOrder.Market,
        '    '                            .Supporting1 = "Force Entry"
        '    '                        }
        '    'ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
        '    '    parameter = New PlaceOrderParameters With {
        '    '                    .EntryPrice = currentTick.Open,
        '    '                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
        '    '                    .Quantity = quantity,
        '    '                    .Stoploss = .EntryPrice + 100000,
        '    '                    .Target = .EntryPrice - 100000,
        '    '                    .Buffer = 0,
        '    '                    .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
        '    '                    .OrderType = Trade.TypeOfOrder.Market,
        '    '                    .Supporting1 = "Force Entry"
        '    '                }
        '    'End If

        '    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        '    Me.ForceTakeTrade = False
        'Else
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
            End If
            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromInvestment(Me.LotSize, 10000, currentTick.Open, _parentStrategy.StockType, True)

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
                                    .Supporting1 = _lastSwingHigh.ToString("HH:mm:ss")
                                }
                    CType(Me.AnotherPairInstrument, PairStrategyRule).Direction = Trade.TradeExecutionDirection.Sell
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
                                    .Supporting1 = _lastSwingLow.ToString("HH:mm:ss")
                                }
                    CType(Me.AnotherPairInstrument, PairStrategyRule).Direction = Trade.TradeExecutionDirection.Buy
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        'End If
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

    Private _lastSwingHigh As Date = Date.MinValue
    Private _lastSwingLow As Date = Date.MinValue
    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection) = Nothing
        If Me.Direction = Trade.TradeExecutionDirection.None OrElse Me.Direction = Trade.TradeExecutionDirection.Buy Then
            If _swingHighPayload(candle.PayloadDate) <> _swingHighPayload(candle.PreviousCandlePayload.PayloadDate) Then
                _lastSwingHigh = candle.PreviousCandlePayload.PayloadDate
            End If
            If _lastSwingHigh <> Date.MinValue AndAlso Me.ChangePercentagePayloads(candle.PayloadDate).High >= Me.ChangePercentagePayloads(_lastSwingHigh).High Then
                If Me.Direction = Trade.TradeExecutionDirection.None Then
                    Dim myPair As PairStrategyRule = Me.AnotherPairInstrument
                    If Me.ChangePercentagePayloads(candle.PayloadDate).Low < myPair.ChangePercentagePayloads(candle.PayloadDate).Low Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Buy)
                    End If
                Else
                    ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Buy)
                End If
            End If
        End If
        If ret Is Nothing AndAlso (Me.Direction = Trade.TradeExecutionDirection.None OrElse Me.Direction = Trade.TradeExecutionDirection.Sell) Then
            If _swingLowPayload(candle.PayloadDate) <> _swingLowPayload(candle.PreviousCandlePayload.PayloadDate) Then
                _lastSwingLow = candle.PreviousCandlePayload.PayloadDate
            End If
            If _lastSwingLow <> Date.MinValue AndAlso Me.ChangePercentagePayloads(candle.PayloadDate).Low <= Me.ChangePercentagePayloads(_lastSwingLow).Low Then
                If Me.Direction = Trade.TradeExecutionDirection.None Then
                    Dim myPair As PairStrategyRule = Me.AnotherPairInstrument
                    If Me.ChangePercentagePayloads(candle.PayloadDate).High > myPair.ChangePercentagePayloads(candle.PayloadDate).High Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Sell)
                    End If
                Else
                    ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function
End Class

Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PairDifferenceStrategyRule
    Inherits StrategyRule

    Public Direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None

    Public PreviousDayFirstCandleClose As Decimal = Decimal.MinValue
    Public LastTick As Payload = Nothing

    Private ReadOnly _controller As Boolean

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, canceller)

        If _tradingSymbol.ToUpper = "MARUTI" Then
            _controller = True
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            Dim previousTradingDay As Date = _parentStrategy.Cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Cash, _tradingDate)

            PreviousDayFirstCandleClose = _signalPayload.Where(Function(x)
                                                                   Return x.Key.Date = previousTradingDay.Date
                                                               End Function).OrderBy(Function(y)
                                                                                         Return y.Key
                                                                                     End Function).FirstOrDefault.Value.Close

        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If Me.ForceTakeTrade AndAlso Not _controller Then
            Dim quantity As Integer = _parentStrategy.CalculateQuantityFromInvestment(Me.LotSize, 10000, currentTick.Open, _parentStrategy.StockType, True)
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
            ElseIf Me.direction = Trade.TradeExecutionDirection.Sell Then
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
                Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
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
                                        .Supporting1 = "Normal Entry"
                                    }
                        CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).Direction = Trade.TradeExecutionDirection.Sell
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
                        CType(Me.AnotherPairInstrument, PairChangePercentStrategyRule).Direction = Trade.TradeExecutionDirection.Buy
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
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
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
            Dim myPairChange As Decimal = ((myPair.LastTick.Open / myPair.PreviousDayFirstCandleClose) - 1) * 100
            Dim myChange As Decimal = ((Me.LastTick.Open / Me.PreviousDayFirstCandleClose) - 1) * 100
            If Math.Abs(myChange - myPairChange) >= 2 Then
                Console.WriteLine(String.Format("{0},{1},{2},{3}", LastTick.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss"), Me.LastTick.Open, myPair.LastTick.Open, Math.Round(Math.Abs(myChange - myPairChange), 2)))
            End If
        End If
        Return ret
    End Function
End Class

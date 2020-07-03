Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PairDifferencePercentageStrategyRule
    Inherits StrategyRule

    Public DummyCandle As Payload = Nothing
    Public Direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Public ChangePercentage As Decimal = Decimal.MinValue
    Public PreviousDayLastCandleClose As Decimal = Decimal.MinValue

    Private _hkPayloads As Dictionary(Of Date, Payload) = Nothing

    Private ReadOnly _tradeStartTime As Date = Date.MinValue
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
        _tradeStartTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        If controller = 1 Then
            _controller = True
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayloads)

            Dim firstCandleOfTheDay As Payload = _signalPayload.Where(Function(x)
                                                                          Return x.Key.Date = _tradingDate.Date
                                                                      End Function).OrderBy(Function(y)
                                                                                                Return y.Key
                                                                                            End Function).FirstOrDefault.Value

            Me.PreviousDayLastCandleClose = firstCandleOfTheDay.PreviousCandlePayload.Close
            Me.ChangePercentage = ((firstCandleOfTheDay.Close / Me.PreviousDayLastCandleClose) - 1) * 100
        End If
    End Sub

    Public Overrides Sub CompletePairProcessing()
        MyBase.CompletePairProcessing()

        If _controller Then
            Dim myPair As PairDifferencePercentageStrategyRule = Me.AnotherPairInstrument
            If Me.ChangePercentage <> Decimal.MinValue AndAlso myPair.ChangePercentage <> Decimal.MinValue Then
                Dim myChange As Decimal = Me.ChangePercentage
                Dim myPairChange As Decimal = myPair.ChangePercentage
                If myChange > myPairChange Then
                    Me.Direction = Trade.TradeExecutionDirection.Sell
                    myPair.Direction = Trade.TradeExecutionDirection.Buy
                Else
                    myPair.Direction = Trade.TradeExecutionDirection.Sell
                    Me.Direction = Trade.TradeExecutionDirection.Buy
                End If
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Me.DummyCandle = currentTick
        If _controller Then

        End If
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item3
            End If
            If signalCandle IsNot Nothing Then
                Dim quantity As Integer = 1
                If currentTick.Open > CType(Me.AnotherPairInstrument, PairDifferencePercentageStrategyRule).DummyCandle.Open Then
                    quantity = Me.LotSize
                Else
                    Dim multiplier As Decimal = CType(Me.AnotherPairInstrument, PairDifferencePercentageStrategyRule).DummyCandle.Open / currentTick.Open
                    quantity = Math.Floor(multiplier * CType(Me.AnotherPairInstrument, PairDifferencePercentageStrategyRule).LotSize)
                End If

                If Me.Direction = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = signal.Item2,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - 100000,
                                    .Target = .EntryPrice + 100000,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = Math.Round(Me.ChangePercentage, 4)
                                }
                ElseIf Me.Direction = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = signal.Item2,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + 100000,
                                    .Target = .EntryPrice - 100000,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = Math.Round(Me.ChangePercentage, 4)
                                }
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
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Payload) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Payload) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayloads(candle.PreviousCandlePayload.PayloadDate)
            If Me.Direction = Trade.TradeExecutionDirection.Buy AndAlso hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                Dim entryPrice As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                If entryPrice = hkCandle.High Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(entryPrice, RoundOfType.Floor)
                    entryPrice = entryPrice + buffer
                End If
                ret = New Tuple(Of Boolean, Decimal, Payload)(True, entryPrice, hkCandle)
            ElseIf Me.Direction = Trade.TradeExecutionDirection.Sell AndAlso hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                Dim entryPrice As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                If entryPrice = hkCandle.Low Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(entryPrice, RoundOfType.Floor)
                    entryPrice = entryPrice - buffer
                End If
                ret = New Tuple(Of Boolean, Decimal, Payload)(True, entryPrice, hkCandle)
            End If
        End If
        Return ret
    End Function
End Class
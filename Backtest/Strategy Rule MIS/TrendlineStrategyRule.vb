Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class TrendlineStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public BuyStoplossLevel As Decimal
        Public SellStoplossLevel As Decimal
        Public MaxLossPerDay As Decimal
        Public MaxTradePerDay As Decimal
    End Class
#End Region

    Private _highTrendLine As Dictionary(Of Date, Decimal) = Nothing
    Private _lowTrendLine As Dictionary(Of Date, Decimal) = Nothing

    Private _buyStoplossPrice As Decimal = Decimal.MinValue
    Private _sellStoplossPrice As Decimal = Decimal.MinValue

    Private ReadOnly _maxlossPerTrade As Decimal
    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities

        _maxlossPerTrade = _userInputs.MaxLossPerDay / _userInputs.MaxTradePerDay
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Dim currentDayPayload As Dictionary(Of Date, Payload) = Nothing
        For Each runningPaylod In _signalPayload
            If runningPaylod.Key.Date = _tradingDate.Date Then
                _cts.Token.ThrowIfCancellationRequested()
                If currentDayPayload Is Nothing Then currentDayPayload = New Dictionary(Of Date, Payload)
                currentDayPayload.Add(runningPaylod.Key, runningPaylod.Value)
            End If
        Next
        If currentDayPayload IsNot Nothing AndAlso currentDayPayload.Count > 0 Then
            Dim firstCandleOfTheDay As Payload = currentDayPayload.FirstOrDefault.Value
            Dim rangePer As Decimal = Math.Round(firstCandleOfTheDay.CandleRange * 61.8 / 100, 2)

            Dim x1 As Decimal = 0
            Dim y1 As Decimal = firstCandleOfTheDay.High + rangePer
            Dim x2 As Decimal = currentDayPayload.Count - 1
            Dim y2 As Decimal = firstCandleOfTheDay.Low + rangePer
            Dim highTrendLine As TrendLineVeriables = Common.GetEquationOfTrendLine(x1, y1, x2, y2)

            Dim p1 As Decimal = 0
            Dim q1 As Decimal = firstCandleOfTheDay.Low - rangePer
            Dim p2 As Decimal = currentDayPayload.Count - 1
            Dim q2 As Decimal = firstCandleOfTheDay.High - rangePer
            Dim lowTrendLine As TrendLineVeriables = Common.GetEquationOfTrendLine(p1, q1, p2, q2)

            Dim counter As Integer = 0
            For Each runningCandle In currentDayPayload
                _cts.Token.ThrowIfCancellationRequested()

                If _highTrendLine Is Nothing Then _highTrendLine = New Dictionary(Of Date, Decimal)
                _highTrendLine.Add(runningCandle.Key, Math.Round(highTrendLine.M * counter + highTrendLine.C, 2))

                If _lowTrendLine Is Nothing Then _lowTrendLine = New Dictionary(Of Date, Decimal)
                _lowTrendLine.Add(runningCandle.Key, Math.Round(lowTrendLine.M * counter + lowTrendLine.C, 2))

                counter += 1
            Next

            _buyStoplossPrice = firstCandleOfTheDay.Low + Math.Round(firstCandleOfTheDay.CandleRange * _userInputs.BuyStoplossLevel / 100, 2)
            _sellStoplossPrice = firstCandleOfTheDay.Low + Math.Round(firstCandleOfTheDay.CandleRange * _userInputs.SellStoplossLevel / 100, 2)
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim totalTrades As Integer = _parentStrategy.TotalNumberOfOpenExecutedTrades(_tradingDate)
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            totalTrades < _userInputs.MaxTradePerDay AndAlso currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2.High, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2.High + buffer
                    Dim stoploss As Decimal = ConvertFloorCeling(_buyStoplossPrice, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim target As Decimal = entryPrice + ConvertFloorCeling((entryPrice - stoploss) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, Math.Abs(_maxlossPerTrade) * -1, Trade.TypeOfStock.Cash)

                    If quantity <> 0 AndAlso currentTick.Open < entryPrice Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signal.Item2,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _highTrendLine(signal.Item2.PayloadDate),
                                    .Supporting3 = GetLastTrendlineTouchCandle(signal.Item2, Trade.TradeExecutionDirection.Buy).PayloadDate.ToString("HH:mm:ss")
                                }
                    End If
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2.Low, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2.Low - buffer
                    Dim stoploss As Decimal = ConvertFloorCeling(_sellStoplossPrice, _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim target As Decimal = entryPrice - ConvertFloorCeling((stoploss - entryPrice) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, Math.Abs(_maxlossPerTrade) * -1, Trade.TypeOfStock.Cash)

                    If quantity <> 0 AndAlso currentTick.Open > entryPrice Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signal.Item2,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _lowTrendLine(signal.Item2.PayloadDate),
                                    .Supporting3 = GetLastTrendlineTouchCandle(signal.Item2, Trade.TradeExecutionDirection.Sell).PayloadDate.ToString("HH:mm:ss")
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
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim highTrendline As Decimal = _highTrendLine(candle.PayloadDate)
            Dim lowTrendline As Decimal = _lowTrendLine(candle.PayloadDate)
            If candle.Low > highTrendline Then
                Dim lastCandle As Payload = GetLastTrendlineTouchCandle(candle, Trade.TradeExecutionDirection.Buy)
                If lastCandle IsNot Nothing AndAlso candle.Close >= lastCandle.High Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, candle, Trade.TradeExecutionDirection.Buy)
                End If
            ElseIf candle.High < lowTrendline Then
                Dim lastCandle As Payload = GetLastTrendlineTouchCandle(candle, Trade.TradeExecutionDirection.Sell)
                If lastCandle IsNot Nothing AndAlso candle.Close <= lastCandle.Low Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, candle, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetLastTrendlineTouchCandle(ByVal signalCandle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Payload
        Dim ret As Payload = Nothing
        For Each runningPayload In _signalPayload.OrderByDescending(Function(x)
                                                                        Return x.Key
                                                                    End Function)
            If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key <= signalCandle.PayloadDate Then
                If direction = Trade.TradeExecutionDirection.Buy Then
                    Dim trendline As Decimal = _highTrendLine(runningPayload.Key)
                    If runningPayload.Value.Low < trendline Then
                        ret = runningPayload.Value
                        Exit For
                    End If
                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                    Dim trendline As Decimal = _lowTrendLine(runningPayload.Key)
                    If runningPayload.Value.High > trendline Then
                        ret = runningPayload.Value
                        Exit For
                    End If
                End If
            End If
        Next
        Return ret
    End Function

End Class
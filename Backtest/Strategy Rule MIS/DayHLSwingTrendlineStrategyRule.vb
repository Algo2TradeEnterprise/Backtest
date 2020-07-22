Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class DayHLSwingTrendlineStrategyRule
    Inherits StrategyRule

    Private _emaPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _highTrendlinePayload As Dictionary(Of Date, TrendLineVeriables) = Nothing
    Private _lowTrendlinePayload As Dictionary(Of Date, TrendLineVeriables) = Nothing
    Private _slPoint As Decimal = Decimal.MinValue
    Private _targetPoint As Decimal = Decimal.MinValue
    Private _quantity As Integer = Integer.MinValue

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.EMA.CalculateEMA(13, Payload.PayloadFields.Close, _signalPayload, _emaPayload)
        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)

        CalculateDayHLSwingTrendline(_signalPayload, _highTrendlinePayload, _lowTrendlinePayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date, Date) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS)
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item2
                    If _quantity = Integer.MinValue Then
                        _slPoint = ConvertFloorCeling(GetHighestATROfTheDay(signalCandle), _parentStrategy.TickSize, RoundOfType.Floor)
                        _quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, currentTick.Open, currentTick.Open - _slPoint, -500, Trade.TypeOfStock.Cash)
                        _targetPoint = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, currentTick.Open, _quantity, 500, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - currentTick.Open
                    End If
                Else
                    If lastExecutedOrder.EntryDirection <> signal.Item3 Then
                        signalCandle = signal.Item2
                    End If
                End If
            End If

                If signalCandle IsNot Nothing AndAlso _quantity <> 0 Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim stoploss As Decimal = entryPrice - 1000000
                    Dim target As Decimal = entryPrice + 1000000
                    Dim quantity As Integer = _quantity

                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = entryPrice + _targetPoint,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = signal.Item4.ToString("HH:mm:ss"),
                                .Supporting3 = signal.Item5.ToString("HH:mm:ss")
                            }

                    parameter2 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = signal.Item4.ToString("HH:mm:ss"),
                                .Supporting3 = signal.Item5.ToString("HH:mm:ss")
                            }
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim stoploss As Decimal = entryPrice + 1000000
                    Dim target As Decimal = entryPrice - 1000000
                    Dim quantity As Integer = _quantity

                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = entryPrice - _targetPoint,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = signal.Item4.ToString("HH:mm:ss"),
                                .Supporting3 = signal.Item5.ToString("HH:mm:ss")
                            }

                    parameter2 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = signal.Item4.ToString("HH:mm:ss"),
                                .Supporting3 = signal.Item5.ToString("HH:mm:ss")
                            }
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter1)
        End If
        If parameter2 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter2)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date, Date) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryDirection <> signal.Item3 Then
                    ret = New Tuple(Of Boolean, String)(True, "Opposite Direction Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim plPoint As Decimal = currentTick.Open - currentTrade.EntryPrice
                If plPoint > _targetPoint Then
                    Dim multiplier As Integer = Math.Floor(plPoint / _targetPoint)
                    If multiplier = 1 Then
                        Dim brkevnPnt As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                        triggerPrice = currentTrade.EntryPrice + brkevnPnt
                    ElseIf multiplier > 1 Then
                        triggerPrice = currentTrade.EntryPrice + (_targetPoint * (multiplier - 1))
                    End If
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim plPoint As Decimal = currentTrade.EntryPrice - currentTick.Open
                If plPoint > _targetPoint Then
                    Dim multiplier As Integer = Math.Floor(plPoint / _targetPoint)
                    If multiplier = 1 Then
                        Dim brkevnPnt As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                        triggerPrice = currentTrade.EntryPrice - brkevnPnt
                    ElseIf multiplier > 1 Then
                        triggerPrice = currentTrade.EntryPrice - (_targetPoint * (multiplier - 1))
                    End If
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso currentTrade.PotentialStopLoss <> triggerPrice Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, Math.Abs(triggerPrice - currentTrade.EntryPrice))
            End If
        End If
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date, Date)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date, Date) = Nothing
        If candle IsNot Nothing Then
            Dim ema As Decimal = _emaPayload(candle.PayloadDate)
            Dim highTrendline As TrendLineVeriables = _highTrendlinePayload(candle.PayloadDate)
            If highTrendline IsNot Nothing AndAlso candle.Low > highTrendline.CurrentValue AndAlso candle.Close > ema Then
                ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date, Date)(True, candle, Trade.TradeExecutionDirection.Buy, highTrendline.Point1, highTrendline.Point2)
            End If

            Dim lowTrendline As TrendLineVeriables = _lowTrendlinePayload(candle.PayloadDate)
            If lowTrendline IsNot Nothing AndAlso candle.High < lowTrendline.CurrentValue AndAlso candle.Close < ema Then
                ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date, Date)(True, candle, Trade.TradeExecutionDirection.Sell, lowTrendline.Point1, lowTrendline.Point2)
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestATROfTheDay(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        ret = _atrPayload.Max(Function(x)
                                  If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                      Return x.Value
                                  Else
                                      Return Decimal.MinValue
                                  End If
                              End Function)
        Return ret
    End Function

    Private Sub CalculateDayHLSwingTrendline(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef highTrendlinePayload As Dictionary(Of Date, TrendLineVeriables), ByRef lowTrendlinePayload As Dictionary(Of Date, TrendLineVeriables))
        Dim swingPayload As Dictionary(Of Date, Indicator.Swing) = Nothing
        Indicator.SwingHighLow.CalculateSwingHighLow(inputPayload, False, swingPayload)

        Dim emaPayload As Dictionary(Of Date, Decimal) = Nothing
        Indicator.EMA.CalculateEMA(13, Payload.PayloadFields.Close, inputPayload, emaPayload)

        Dim lastHighTrendline As TrendLineVeriables = Nothing
        Dim lastLowTrendline As TrendLineVeriables = Nothing
        For Each runningPaylod In inputPayload
            Dim mainHighTrendline As TrendLineVeriables = Nothing
            Dim mainLowTrendline As TrendLineVeriables = Nothing
            Dim highTrendline As TrendLineVeriables = Nothing
            Dim lowTrendline As TrendLineVeriables = Nothing
            If runningPaylod.Value.PreviousCandlePayload IsNot Nothing AndAlso runningPaylod.Value.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
                runningPaylod.Value.PreviousCandlePayload.PreviousCandlePayload.PayloadDate.Date = runningPaylod.Key.Date Then
                Dim swing As Indicator.Swing = swingPayload(runningPaylod.Value.PreviousCandlePayload.PreviousCandlePayload.PayloadDate)
                If swing.SwingHighTime.Date = runningPaylod.Key.Date Then
                    If swing.SwingHigh > emaPayload(swing.SwingHighTime) Then
                        Dim highestCandle As Payload = GetHighestHigh(inputPayload, runningPaylod.Value.PreviousCandlePayload)
                        If lastHighTrendline IsNot Nothing AndAlso lastHighTrendline.Point1 <> highestCandle.PayloadDate Then lastHighTrendline = Nothing
                        If swing.SwingHighTime > highestCandle.PayloadDate AndAlso swing.SwingHigh < highestCandle.High Then
                            Dim swingTime As Date = swing.SwingHighTime
                            If inputPayload(swingTime.AddMinutes(_parentStrategy.SignalTimeFrame)).High = swing.SwingHigh Then
                                swingTime = swingTime.AddMinutes(_parentStrategy.SignalTimeFrame)
                            End If
                            Dim x1 As Decimal = 0
                            Dim y1 As Decimal = highestCandle.High
                            Dim x2 As Decimal = inputPayload.Where(Function(x)
                                                                       Return x.Key > highestCandle.PayloadDate AndAlso x.Key <= swingTime
                                                                   End Function).Count
                            Dim y2 As Decimal = swing.SwingHigh
                            Dim trendLine As TrendLineVeriables = Common.GetEquationOfTrendLine(x1, y1, x2, y2)
                            trendLine.Point1 = highestCandle.PayloadDate
                            trendLine.Point2 = swingTime
                            If trendLine IsNot Nothing AndAlso IsValidTrendLine(trendLine, inputPayload) Then
                                highTrendline = trendLine
                            End If
                        End If
                    End If
                End If
                If swing.SwingLowTime.Date = runningPaylod.Key.Date Then
                    If swing.SwingLow < emaPayload(swing.SwingLowTime) Then
                        Dim lowestCandle As Payload = GetLowestLow(inputPayload, runningPaylod.Value.PreviousCandlePayload)
                        If lastLowTrendline IsNot Nothing AndAlso lastLowTrendline.Point1 <> lowestCandle.PayloadDate Then lastLowTrendline = Nothing
                        If swing.SwingLowTime > lowestCandle.PayloadDate AndAlso swing.SwingLow > lowestCandle.Low Then
                            Dim swingTime As Date = swing.SwingLowTime
                            If inputPayload(swingTime.AddMinutes(_parentStrategy.SignalTimeFrame)).Low = swing.SwingLow Then
                                swingTime = swingTime.AddMinutes(_parentStrategy.SignalTimeFrame)
                            End If
                            Dim x1 As Decimal = 0
                            Dim y1 As Decimal = lowestCandle.Low
                            Dim x2 As Decimal = inputPayload.Where(Function(x)
                                                                       Return x.Key > lowestCandle.PayloadDate AndAlso x.Key <= swingTime
                                                                   End Function).Count
                            Dim y2 As Decimal = swing.SwingLow
                            Dim trendLine As TrendLineVeriables = Common.GetEquationOfTrendLine(x1, y1, x2, y2)
                            trendLine.Point1 = lowestCandle.PayloadDate
                            trendLine.Point2 = swingTime
                            If trendLine IsNot Nothing AndAlso IsValidTrendLine(trendLine, inputPayload) Then
                                lowTrendline = trendLine
                            End If
                        End If
                    End If
                End If
            End If

            If highTrendline Is Nothing Then
                If lastHighTrendline IsNot Nothing AndAlso lastHighTrendline.Point1.Date = runningPaylod.Key.Date Then
                    highTrendline = lastHighTrendline
                End If
            End If
            If highTrendline IsNot Nothing Then
                Dim counter As Integer = inputPayload.Where(Function(x)
                                                                Return x.Key > highTrendline.Point1 AndAlso x.Key <= runningPaylod.Key
                                                            End Function).Count

                mainHighTrendline = New TrendLineVeriables With {
                    .C = highTrendline.C,
                    .M = highTrendline.M,
                    .X = counter,
                    .Point1 = highTrendline.Point1,
                    .Point2 = highTrendline.Point2
                }
            End If
            If lowTrendline Is Nothing Then
                If lastLowTrendline IsNot Nothing AndAlso lastLowTrendline.Point1.Date = runningPaylod.Key.Date Then
                    lowTrendline = lastLowTrendline
                End If
            End If
            If lowTrendline IsNot Nothing Then
                Dim counter As Integer = inputPayload.Where(Function(x)
                                                                Return x.Key > lowTrendline.Point1 AndAlso x.Key <= runningPaylod.Key
                                                            End Function).Count
                mainLowTrendline = New TrendLineVeriables With {
                    .C = lowTrendline.C,
                    .M = lowTrendline.M,
                    .X = counter,
                    .Point1 = lowTrendline.Point1,
                    .Point2 = lowTrendline.Point2
                }
            End If

            If highTrendlinePayload Is Nothing Then highTrendlinePayload = New Dictionary(Of Date, TrendLineVeriables)
            highTrendlinePayload.Add(runningPaylod.Key, mainHighTrendline)
            If lowTrendlinePayload Is Nothing Then lowTrendlinePayload = New Dictionary(Of Date, TrendLineVeriables)
            lowTrendlinePayload.Add(runningPaylod.Key, mainLowTrendline)

            lastHighTrendline = mainHighTrendline
            lastLowTrendline = mainLowTrendline
        Next
    End Sub

    Private Function IsValidTrendLine(ByVal trendline As TrendLineVeriables, ByVal inputPayload As Dictionary(Of Date, Payload)) As Boolean
        Dim ret As Boolean = True
        For Each runningPayload In inputPayload
            If runningPayload.Key > trendline.Point1 AndAlso runningPayload.Key < trendline.Point2 Then
                Dim counter As Integer = inputPayload.Where(Function(x)
                                                                Return x.Key > trendline.Point1 AndAlso x.Key <= runningPayload.Key
                                                            End Function).Count
                Dim point As Decimal = Math.Round(trendline.M * counter + trendline.C, 2)
                If runningPayload.Value.High > point AndAlso runningPayload.Value.Low < point Then
                    ret = False
                    Exit For
                End If
            End If
        Next
        Return ret
    End Function

    Private Function GetHighestHigh(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal signalCandle As Payload) As Payload
        Dim ret As Payload = Nothing
        For Each runningPayload In inputPayload
            If runningPayload.Key.Date = signalCandle.PayloadDate.Date AndAlso runningPayload.Key <= signalCandle.PayloadDate Then
                If ret Is Nothing Then ret = runningPayload.Value
                If runningPayload.Value.High >= ret.High Then
                    ret = runningPayload.Value
                End If
            End If
        Next
        Return ret
    End Function

    Private Function GetLowestLow(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal signalCandle As Payload) As Payload
        Dim ret As Payload = Nothing
        For Each runningPayload In inputPayload
            If runningPayload.Key.Date = signalCandle.PayloadDate.Date AndAlso runningPayload.Key <= signalCandle.PayloadDate Then
                If ret Is Nothing Then ret = runningPayload.Value
                If runningPayload.Value.Low <= ret.Low Then
                    ret = runningPayload.Value
                End If
            End If
        Next
        Return ret
    End Function
End Class
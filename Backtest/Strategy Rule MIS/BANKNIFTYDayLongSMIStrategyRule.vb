Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class BANKNIFTYDayLongSMIStrategyRule
    Inherits StrategyRule

    Private _SwingHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _SwingLowPayload As Dictionary(Of Date, Decimal) = Nothing

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

        Dim SMISignalPayload As Dictionary(Of Date, Decimal) = Nothing
        Dim EMASignalPayload As Dictionary(Of Date, Decimal) = Nothing
        Indicator.SMI.CalculateSMI(10, 3, 3, 10, _signalPayload, SMISignalPayload, EMASignalPayload)
        Dim SMISignalInputPayload As Dictionary(Of Date, Payload) = Nothing
        ConvertDecimalToPayload(SMISignalPayload, SMISignalInputPayload)
        Indicator.SwingHighLow.CalculateSwingHighLow(SMISignalInputPayload, False, _SwingHighPayload, _SwingLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > Math.Abs(_parentStrategy.OverAllLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime Then
            Dim longTrades As List(Of Trade) = _parentStrategy.GetOpenActiveTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy)
            Dim shortTrades As List(Of Trade) = _parentStrategy.GetOpenActiveTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell)

            If longTrades Is Nothing OrElse (longTrades IsNot Nothing AndAlso longTrades.Count = 0) Then
                Dim currentSignal As Tuple(Of Boolean, Decimal, Decimal, Payload) = GetSignals(currentMinuteCandlePayload.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy)
                If currentSignal IsNot Nothing AndAlso currentSignal.Item1 Then
                    'If currentTick.Open < currentSignal.Item2 Then
                    'Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentSignal.Item2, RoundOfType.Floor)
                    Dim buffer As Decimal = 2
                    parameter1 = New PlaceOrderParameters With {
                        .EntryPrice = currentSignal.Item2 + buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = LotSize * 10,
                        .Stoploss = currentSignal.Item3 - buffer,
                        .Target = .EntryPrice + 100000,
                        .Buffer = buffer,
                        .SignalCandle = currentSignal.Item4,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = currentSignal.Item4.PayloadDate.ToShortTimeString
                    }
                    'End If
                End If
            End If
            If shortTrades Is Nothing OrElse (shortTrades IsNot Nothing AndAlso shortTrades.Count = 0) Then
                Dim currentSignal As Tuple(Of Boolean, Decimal, Decimal, Payload) = GetSignals(currentMinuteCandlePayload.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell)
                If currentSignal IsNot Nothing AndAlso currentSignal.Item1 Then
                    'If currentTick.Open > currentSignal.Item3 Then
                    'Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentSignal.Item3, RoundOfType.Floor)
                    Dim buffer As Decimal = 2
                    parameter2 = New PlaceOrderParameters With {
                        .EntryPrice = currentSignal.Item3 - buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                        .Quantity = LotSize * 10,
                        .Stoploss = currentSignal.Item2 + buffer,
                        .Target = .EntryPrice - 100000,
                        .Buffer = buffer,
                        .SignalCandle = currentSignal.Item4,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = currentSignal.Item4.PayloadDate.ToShortTimeString
                    }
                    'End If
                End If
            End If
        End If
        Dim orderList As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            If orderList Is Nothing Then orderList = New List(Of PlaceOrderParameters)
            orderList.Add(parameter1)
        End If
        If parameter2 IsNot Nothing Then
            If orderList Is Nothing Then orderList = New List(Of PlaceOrderParameters)
            orderList.Add(parameter2)
        End If
        If orderList IsNot Nothing AndAlso orderList.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, orderList)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentSignal As Tuple(Of Boolean, Decimal, Decimal, Payload) = GetSignals(currentMinuteCandlePayload.PreviousCandlePayload, currentTrade.EntryDirection)
            If currentSignal IsNot Nothing AndAlso currentSignal.Item1 Then
                Dim entryPrice As Decimal = Decimal.MinValue
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    entryPrice = currentSignal.Item2 + 2
                    If entryPrice < currentTrade.EntryPrice Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    entryPrice = currentSignal.Item3 - 2
                    If entryPrice > currentTrade.EntryPrice Then
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                direction = Trade.TradeExecutionDirection.Sell
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                direction = Trade.TradeExecutionDirection.Buy
            End If
            Dim currentTrades As List(Of Trade) = _parentStrategy.GetOpenActiveTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, direction)
            If currentTrades IsNot Nothing AndAlso currentTrades.Count > 0 Then
                Dim triggerPrice As Decimal = currentTrades.FirstOrDefault.EntryPrice
                If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("{0}. Time:{1}", triggerPrice, currentTick.PayloadDate))
                End If
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

    'Private Function GetSignals(ByVal candle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Decimal, Decimal)
    '    Dim ret As Tuple(Of Boolean, Decimal, Decimal) = Nothing
    '    If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
    '        If direction = Trade.TradeExecutionDirection.Buy Then
    '            Dim previousSwingHigh As Decimal = _SwingHighPayload(candle.PreviousCandlePayload.PayloadDate)
    '            Dim currentSwingHigh As Decimal = _SwingHighPayload(candle.PayloadDate)
    '            If previousSwingHigh <> currentSwingHigh Then
    '                Dim highestHigh As Decimal = Math.Max(Math.Max(candle.High, candle.PreviousCandlePayload.High), candle.PreviousCandlePayload.PreviousCandlePayload.High)
    '                Dim lowestLow As Decimal = Math.Min(Math.Min(candle.Low, candle.PreviousCandlePayload.Low), candle.PreviousCandlePayload.PreviousCandlePayload.Low)
    '                Dim dummyHighestHigh As Decimal = Decimal.MinValue
    '                If highestHigh = candle.PreviousCandlePayload.PreviousCandlePayload.High Then
    '                    dummyHighestHigh = Math.Max(candle.High, candle.PreviousCandlePayload.High)
    '                ElseIf highestHigh = candle.PreviousCandlePayload.High Then
    '                    dummyHighestHigh = Math.Max(candle.High, candle.PreviousCandlePayload.PreviousCandlePayload.High)
    '                ElseIf highestHigh = candle.High Then
    '                    dummyHighestHigh = Math.Max(candle.PreviousCandlePayload.PreviousCandlePayload.High, candle.PreviousCandlePayload.High)
    '                End If
    '                If dummyHighestHigh < (highestHigh + lowestLow) / 2 Then
    '                    highestHigh = dummyHighestHigh
    '                End If
    '                ret = New Tuple(Of Boolean, Decimal, Decimal)(True, highestHigh, lowestLow)
    '            End If
    '        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
    '            Dim previousSwingLow As Decimal = _SwingLowPayload(candle.PreviousCandlePayload.PayloadDate)
    '            Dim currentSwingLow As Decimal = _SwingLowPayload(candle.PayloadDate)
    '            If previousSwingLow <> currentSwingLow Then
    '                Dim highestHigh As Decimal = Math.Max(Math.Max(candle.High, candle.PreviousCandlePayload.High), candle.PreviousCandlePayload.PreviousCandlePayload.High)
    '                Dim lowestLow As Decimal = Math.Min(Math.Min(candle.Low, candle.PreviousCandlePayload.Low), candle.PreviousCandlePayload.PreviousCandlePayload.Low)
    '                Dim dummyLowestLow As Decimal = Decimal.MinValue
    '                If lowestLow = candle.PreviousCandlePayload.PreviousCandlePayload.Low Then
    '                    dummyLowestLow = Math.Min(candle.Low, candle.PreviousCandlePayload.Low)
    '                ElseIf lowestLow = candle.PreviousCandlePayload.Low Then
    '                    dummyLowestLow = Math.Min(candle.Low, candle.PreviousCandlePayload.PreviousCandlePayload.Low)
    '                ElseIf lowestLow = candle.Low Then
    '                    dummyLowestLow = Math.Min(candle.PreviousCandlePayload.PreviousCandlePayload.Low, candle.PreviousCandlePayload.Low)
    '                End If
    '                If dummyLowestLow > (highestHigh + lowestLow) / 2 Then
    '                    lowestLow = dummyLowestLow
    '                End If
    '                ret = New Tuple(Of Boolean, Decimal, Decimal)(True, highestHigh, lowestLow)
    '            End If
    '        End If
    '    End If
    '    Return ret
    'End Function

    Private Function GetSignals(ByVal candle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Decimal, Decimal, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Payload) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If direction = Trade.TradeExecutionDirection.Buy Then
                Dim currentSwingHigh As Decimal = _SwingHighPayload(candle.PayloadDate)
                Dim startTime As Date = GetStartTimeOfIndicator(candle.PayloadDate, _SwingHighPayload)
                If startTime.Date = _tradingDate.Date Then
                    candle = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(startTime, _signalPayload))

                    Dim highestHigh As Decimal = Math.Max(Math.Max(candle.High, candle.PreviousCandlePayload.High), candle.PreviousCandlePayload.PreviousCandlePayload.High)
                    Dim lowestLow As Decimal = Math.Min(Math.Min(candle.Low, candle.PreviousCandlePayload.Low), candle.PreviousCandlePayload.PreviousCandlePayload.Low)
                    Dim dummyHighestHigh As Decimal = Decimal.MinValue
                    If highestHigh = candle.PreviousCandlePayload.PreviousCandlePayload.High Then
                        dummyHighestHigh = Math.Max(candle.High, candle.PreviousCandlePayload.High)
                    ElseIf highestHigh = candle.PreviousCandlePayload.High Then
                        dummyHighestHigh = Math.Max(candle.High, candle.PreviousCandlePayload.PreviousCandlePayload.High)
                    ElseIf highestHigh = candle.High Then
                        dummyHighestHigh = Math.Max(candle.PreviousCandlePayload.PreviousCandlePayload.High, candle.PreviousCandlePayload.High)
                    End If
                    If dummyHighestHigh < (highestHigh + lowestLow) / 2 Then
                        highestHigh = dummyHighestHigh
                    End If
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload)(True, highestHigh, lowestLow, candle)
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                Dim currentSwingLow As Decimal = _SwingLowPayload(candle.PayloadDate)
                Dim startTime As Date = GetStartTimeOfIndicator(candle.PayloadDate, _SwingLowPayload)
                If startTime.Date = _tradingDate.Date Then
                    candle = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(startTime, _signalPayload))

                    Dim highestHigh As Decimal = Math.Max(Math.Max(candle.High, candle.PreviousCandlePayload.High), candle.PreviousCandlePayload.PreviousCandlePayload.High)
                    Dim lowestLow As Decimal = Math.Min(Math.Min(candle.Low, candle.PreviousCandlePayload.Low), candle.PreviousCandlePayload.PreviousCandlePayload.Low)
                    Dim dummyLowestLow As Decimal = Decimal.MinValue
                    If lowestLow = candle.PreviousCandlePayload.PreviousCandlePayload.Low Then
                        dummyLowestLow = Math.Min(candle.Low, candle.PreviousCandlePayload.Low)
                    ElseIf lowestLow = candle.PreviousCandlePayload.Low Then
                        dummyLowestLow = Math.Min(candle.Low, candle.PreviousCandlePayload.PreviousCandlePayload.Low)
                    ElseIf lowestLow = candle.Low Then
                        dummyLowestLow = Math.Min(candle.PreviousCandlePayload.PreviousCandlePayload.Low, candle.PreviousCandlePayload.Low)
                    End If
                    If dummyLowestLow > (highestHigh + lowestLow) / 2 Then
                        lowestLow = dummyLowestLow
                    End If
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload)(True, highestHigh, lowestLow, candle)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function ConvertDecimalToPayload(ByVal inputpayload As Dictionary(Of Date, Decimal), ByRef outputpayload As Dictionary(Of Date, Payload))
        Dim previousPayload As Payload = Nothing
        For Each runningitem In inputpayload
            Dim output As Payload = New Payload(Payload.CandleDataSource.Chart)
            output.PayloadDate = runningitem.Key
            output.Open = runningitem.Value
            output.Low = runningitem.Value
            output.High = runningitem.Value
            output.Close = runningitem.Value
            output.Volume = 1
            output.PreviousCandlePayload = previousPayload
            If outputpayload Is Nothing Then outputpayload = New Dictionary(Of Date, Payload)
            outputpayload.Add(runningitem.Key, output)
            previousPayload = output
        Next
        Return Nothing
    End Function

    Private Function GetStartTimeOfIndicator(ByVal currentTime As Date, ByVal indicatorPayload As Dictionary(Of Date, Decimal)) As Date
        Dim ret As Date = Date.MinValue
        Dim currentIndicatorValue As Decimal = indicatorPayload(currentTime)
        Dim indicatorStartTime As Date = Date.MinValue
        For Each runningIndicator In indicatorPayload.Keys.OrderByDescending(Function(x)
                                                                                 Return x
                                                                             End Function)
            If runningIndicator <= currentTime Then
                If indicatorPayload(runningIndicator) <> currentIndicatorValue Then
                    indicatorStartTime = runningIndicator
                    Exit For
                End If
            End If
        Next
        If indicatorStartTime <> Date.MinValue Then
            ret = indicatorStartTime.AddMinutes(_parentStrategy.SignalTimeFrame)
        End If
        Return ret
    End Function
End Class

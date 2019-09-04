Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class SMIFractalStrategyRule
    Inherits StrategyRule

    Private _SMISignalPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _EMASignalPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _FractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _FractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _SwingHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _SwingLowPayload As Dictionary(Of Date, Decimal) = Nothing

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.SMI.CalculateSMI(10, 3, 3, 10, _signalPayload, _SMISignalPayload, _EMASignalPayload)
        Indicator.FractalBands.CalculateFractal(_signalPayload, _FractalHighPayload, _FractalLowPayload)
        Indicator.SwingHighLow.CalculateSwingHighLow(_signalPayload, True, _SwingHighPayload, _SwingLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TradeType.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TradeType.MIS) AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > Math.Abs(_parentStrategy.OverAllLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime Then
            Dim signalCandle As Payload = Nothing

            Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
            End If
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TradeType.MIS)
            Dim lastTradeGain As Decimal = 0
            Dim exitCandle As Payload = Nothing
            If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                lastTradeGain = (lastExecutedTrade.PLPoint / lastExecutedTrade.EntryPrice) * 100
                exitCandle = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedTrade.ExitTime, _signalPayload))
            End If
            If lastTradeGain < 1 AndAlso signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy Then
                    Dim indicatorPayload As Dictionary(Of Date, Decimal) = Nothing
                    If currentTick.Open < _FractalHighPayload(signalCandle.PayloadDate) AndAlso currentTick.Open < _SwingHighPayload(signalCandle.PayloadDate) Then
                        If _SwingHighPayload(signalCandle.PayloadDate) < _FractalHighPayload(signalCandle.PayloadDate) Then
                            indicatorPayload = _SwingHighPayload
                            If exitCandle IsNot Nothing AndAlso currentMinuteCandlePayload.PayloadDate < exitCandle.PayloadDate.AddMinutes(3 * _parentStrategy.SignalTimeFrame) Then
                                indicatorPayload = _FractalHighPayload
                            End If
                        Else
                            indicatorPayload = _FractalHighPayload
                        End If
                    ElseIf currentTick.Open < _FractalHighPayload(signalCandle.PayloadDate) Then
                        indicatorPayload = _FractalHighPayload
                    ElseIf currentTick.Open < _SwingHighPayload(signalCandle.PayloadDate) Then
                        indicatorPayload = _SwingHighPayload
                        If exitCandle IsNot Nothing AndAlso currentMinuteCandlePayload.PayloadDate < exitCandle.PayloadDate.AddMinutes(3 * _parentStrategy.SignalTimeFrame) Then
                            indicatorPayload = _FractalHighPayload
                        End If
                    End If
                    If indicatorPayload IsNot Nothing AndAlso Not IsIndicatorUsed(currentMinuteCandlePayload, indicatorPayload, Trade.TradeExecutionDirection.Buy) Then
                        Dim potentialEntryPrice As Decimal = indicatorPayload(signalCandle.PayloadDate)
                        'If currentTick.Open < potentialEntryPrice Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(potentialEntryPrice, RoundOfType.Floor)
                        Dim entryPrice As Decimal = ConvertFloorCeling(potentialEntryPrice, _parentStrategy.TickSize, RoundOfType.Celing)
                        parameter = New PlaceOrderParameters With {
                            .EntryPrice = entryPrice + buffer,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = _lotSize,
                            .Stoploss = ConvertFloorCeling(_FractalLowPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing) - buffer,
                            .Target = .EntryPrice + 100000,
                            .Buffer = buffer,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.SL,
                            .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                            .Supporting2 = ConvertFloorCeling(_FractalLowPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                        }
                        'End If
                    End If
                ElseIf signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell Then
                    Dim indicatorPayload As Dictionary(Of Date, Decimal) = Nothing
                    If currentTick.Open > _FractalLowPayload(signalCandle.PayloadDate) AndAlso currentTick.Open > _SwingLowPayload(signalCandle.PayloadDate) Then
                        If _SwingLowPayload(signalCandle.PayloadDate) > _FractalLowPayload(signalCandle.PayloadDate) Then
                            indicatorPayload = _SwingLowPayload
                            If exitCandle IsNot Nothing AndAlso currentMinuteCandlePayload.PayloadDate < exitCandle.PayloadDate.AddMinutes(3 * _parentStrategy.SignalTimeFrame) Then
                                indicatorPayload = _FractalLowPayload
                            End If
                        Else
                            indicatorPayload = _FractalLowPayload
                        End If
                    ElseIf currentTick.Open > _FractalLowPayload(signalCandle.PayloadDate) Then
                        indicatorPayload = _FractalLowPayload
                    ElseIf currentTick.Open > _SwingLowPayload(signalCandle.PayloadDate) Then
                        indicatorPayload = _SwingLowPayload
                        If exitCandle IsNot Nothing AndAlso currentMinuteCandlePayload.PayloadDate < exitCandle.PayloadDate.AddMinutes(3 * _parentStrategy.SignalTimeFrame) Then
                            indicatorPayload = _FractalLowPayload
                        End If
                    End If
                    If indicatorPayload IsNot Nothing AndAlso Not IsIndicatorUsed(currentMinuteCandlePayload, indicatorPayload, Trade.TradeExecutionDirection.Sell) Then
                        Dim potentialEntryPrice As Decimal = indicatorPayload(signalCandle.PayloadDate)
                        'If currentTick.Open > potentialEntryPrice Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(potentialEntryPrice, RoundOfType.Floor)
                        Dim entryPrice As Decimal = ConvertFloorCeling(potentialEntryPrice, _parentStrategy.TickSize, RoundOfType.Celing)
                        parameter = New PlaceOrderParameters With {
                            .EntryPrice = entryPrice - buffer,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = _lotSize,
                            .Stoploss = ConvertFloorCeling(_FractalHighPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing) + buffer,
                            .Target = .EntryPrice - 100000,
                            .Buffer = buffer,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.SL,
                            .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                            .Supporting2 = ConvertFloorCeling(_FractalHighPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                        }
                        'End If
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
            Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If currentTrade.EntryDirection <> signalCandleSatisfied.Item2 Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                Else
                    Dim entryPrice As Decimal = Decimal.MinValue
                    Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TradeType.MIS)
                    Dim exitCandle As Payload = Nothing
                    If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                        exitCandle = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedTrade.ExitTime, _signalPayload))
                    End If
                    Dim signalCandle As Payload = currentMinuteCandlePayload.PreviousCandlePayload
                    If signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy Then
                        Dim indicatorPayload As Dictionary(Of Date, Decimal) = Nothing
                        If currentMinuteCandlePayload.Open < _FractalHighPayload(signalCandle.PayloadDate) AndAlso currentMinuteCandlePayload.Open < _SwingHighPayload(signalCandle.PayloadDate) Then
                            If _SwingHighPayload(signalCandle.PayloadDate) < _FractalHighPayload(signalCandle.PayloadDate) Then
                                indicatorPayload = _SwingHighPayload
                                If exitCandle IsNot Nothing AndAlso currentMinuteCandlePayload.PayloadDate < exitCandle.PayloadDate.AddMinutes(3 * _parentStrategy.SignalTimeFrame) Then
                                    indicatorPayload = _FractalHighPayload
                                End If
                            Else
                                indicatorPayload = _FractalHighPayload
                            End If
                        ElseIf currentMinuteCandlePayload.Open < _FractalHighPayload(signalCandle.PayloadDate) Then
                            indicatorPayload = _FractalHighPayload
                        ElseIf currentMinuteCandlePayload.Open < _SwingHighPayload(signalCandle.PayloadDate) Then
                            indicatorPayload = _SwingHighPayload
                            If exitCandle IsNot Nothing AndAlso currentMinuteCandlePayload.PayloadDate < exitCandle.PayloadDate.AddMinutes(3 * _parentStrategy.SignalTimeFrame) Then
                                indicatorPayload = _FractalHighPayload
                            End If
                        End If
                        If indicatorPayload IsNot Nothing Then
                            Dim potentialEntryPrice As Decimal = indicatorPayload(signalCandle.PayloadDate)
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(potentialEntryPrice, RoundOfType.Floor)
                            entryPrice = ConvertFloorCeling(potentialEntryPrice, _parentStrategy.TickSize, RoundOfType.Celing) + buffer
                        End If
                    ElseIf signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell Then
                        Dim indicatorPayload As Dictionary(Of Date, Decimal) = Nothing
                        If currentMinuteCandlePayload.Open > _FractalLowPayload(signalCandle.PayloadDate) AndAlso currentMinuteCandlePayload.Open > _SwingLowPayload(signalCandle.PayloadDate) Then
                            If _SwingLowPayload(signalCandle.PayloadDate) > _FractalLowPayload(signalCandle.PayloadDate) Then
                                indicatorPayload = _SwingLowPayload
                                If exitCandle IsNot Nothing AndAlso currentMinuteCandlePayload.PayloadDate < exitCandle.PayloadDate.AddMinutes(3 * _parentStrategy.SignalTimeFrame) Then
                                    indicatorPayload = _FractalLowPayload
                                End If
                            Else
                                indicatorPayload = _FractalLowPayload
                            End If
                        ElseIf currentMinuteCandlePayload.Open > _FractalLowPayload(signalCandle.PayloadDate) Then
                            indicatorPayload = _FractalLowPayload
                        ElseIf currentMinuteCandlePayload.Open > _SwingLowPayload(signalCandle.PayloadDate) Then
                            indicatorPayload = _SwingLowPayload
                            If exitCandle IsNot Nothing AndAlso currentMinuteCandlePayload.PayloadDate < exitCandle.PayloadDate.AddMinutes(3 * _parentStrategy.SignalTimeFrame) Then
                                indicatorPayload = _FractalLowPayload
                            End If
                        End If
                        If indicatorPayload IsNot Nothing Then
                            Dim potentialEntryPrice As Decimal = indicatorPayload(signalCandle.PayloadDate)
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(potentialEntryPrice, RoundOfType.Floor)
                            entryPrice = ConvertFloorCeling(potentialEntryPrice, _parentStrategy.TickSize, RoundOfType.Celing) - buffer
                        End If
                    End If
                    If entryPrice <> Decimal.MinValue AndAlso entryPrice <> currentTrade.EntryPrice Then
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

        If _parentStrategy.ModifyStoploss AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim remark As String = Nothing
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTrade.EntryPrice, RoundOfType.Floor)
            Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 AndAlso signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell Then
                    triggerPrice = currentMinuteCandlePayload.PreviousCandlePayload.Low - buffer
                    remark = String.Format("Move to candle: {0}. Time:{1}", triggerPrice, currentTick.PayloadDate)
                Else
                    If _FractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) > currentTrade.PotentialStopLoss Then
                        triggerPrice = _FractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) - buffer
                        remark = String.Format("Move to fractal: {0}. Time:{1}", triggerPrice, currentTick.PayloadDate)
                    End If
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 AndAlso signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy Then
                    triggerPrice = currentMinuteCandlePayload.PreviousCandlePayload.High + buffer
                    remark = String.Format("Move to candle: {0}. Time:{1}", triggerPrice, currentTick.PayloadDate)
                Else
                    If _FractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) < currentTrade.PotentialStopLoss Then
                        triggerPrice = _FractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) + buffer
                        remark = String.Format("Move to fractal: {0}. Time:{1}", triggerPrice, currentTick.PayloadDate)
                    End If
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, remark)
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function IsSignalCandle(ByVal candle As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If _SMISignalPayload(candle.PayloadDate) > _EMASignalPayload(candle.PayloadDate) Then
                ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Buy)
            ElseIf _SMISignalPayload(candle.PayloadDate) < _EMASignalPayload(candle.PayloadDate) Then
                ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Sell)
            End If
        End If
        Return ret
    End Function

    Private Function IsIndicatorUsed(ByVal currentTick As Payload, ByVal indicatorPayload As Dictionary(Of Date, Decimal), ByVal entryDirection As Trade.TradeExecutionDirection) As Boolean
        Dim ret As Boolean = False
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim currentIndicatorValue As Decimal = indicatorPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
        Dim currentIndicatorStartTime As Date = Date.MinValue
        For Each runningIndicator In indicatorPayload.Keys.OrderByDescending(Function(x)
                                                                                 Return x
                                                                             End Function)
            If runningIndicator < currentMinuteCandlePayload.PayloadDate Then
                If indicatorPayload(runningIndicator) <> currentIndicatorValue Then
                    currentIndicatorStartTime = runningIndicator
                    Exit For
                End If
            End If
        Next
        If currentIndicatorStartTime <> Date.MinValue Then
            If entryDirection = Trade.TradeExecutionDirection.Buy Then
                currentIndicatorValue += _parentStrategy.CalculateBuffer(currentIndicatorValue, RoundOfType.Floor)
            ElseIf entryDirection = Trade.TradeExecutionDirection.Sell Then
                currentIndicatorValue -= _parentStrategy.CalculateBuffer(currentIndicatorValue, RoundOfType.Floor)
            End If
            ret = IsSignalTriggered(currentIndicatorValue, entryDirection, currentIndicatorStartTime.AddMinutes(_parentStrategy.SignalTimeFrame * 2), currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
        End If
        Return ret
    End Function
End Class

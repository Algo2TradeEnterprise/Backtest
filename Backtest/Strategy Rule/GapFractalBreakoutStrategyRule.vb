Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class GapFractalBreakoutStrategyRule
    Inherits StrategyRule

    Private _FractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _FractalLowPayload As Dictionary(Of Date, Decimal) = Nothing

    Private ReadOnly _Gap As Decimal

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal gap As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller)
        _Gap = gap
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.FractalBands.CalculateFractal(_signalPayload, _FractalHighPayload, _FractalLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrder(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
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

            Dim signalCandleSatisfied As Tuple(Of Boolean, String) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
            End If
            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
                If _Gap > 0.5 Then
                    direction = Trade.TradeExecutionDirection.Sell
                ElseIf _Gap < -0.5 Then
                    direction = Trade.TradeExecutionDirection.Buy
                End If
                If direction = Trade.TradeExecutionDirection.Buy Then
                    Dim potentialEntryPrice As Decimal = _FractalHighPayload(signalCandle.PayloadDate)
                    If currentTick.Open < potentialEntryPrice Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(potentialEntryPrice, RoundOfType.Floor)
                        Dim entryPrice As Decimal = ConvertFloorCeling(potentialEntryPrice, _parentStrategy.TickSize, RoundOfType.Celing)
                        If entryPrice + buffer < GetLastDayLastCandleClose() Then
                            parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice + buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = _lotSize,
                                .Stoploss = ConvertFloorCeling(_FractalLowPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing) - buffer,
                                .Target = .EntryPrice + 100000,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                                .Supporting2 = signalCandleSatisfied.Item2
                            }
                        End If
                    End If
                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                    Dim potentialEntryPrice As Decimal = _FractalLowPayload(signalCandle.PayloadDate)
                    If currentTick.Open > potentialEntryPrice Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(potentialEntryPrice, RoundOfType.Floor)
                        Dim entryPrice As Decimal = ConvertFloorCeling(potentialEntryPrice, _parentStrategy.TickSize, RoundOfType.Celing)
                        If entryPrice - buffer > GetLastDayLastCandleClose() Then
                            parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice - buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = _lotSize,
                                .Stoploss = ConvertFloorCeling(_FractalHighPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing) + buffer,
                                .Target = .EntryPrice - 100000,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                                .Supporting2 = signalCandleSatisfied.Item2
                            }
                        End If
                    End If
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrder(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandleSatisfied As Tuple(Of Boolean, String) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                Dim direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
                If _Gap > 0.5 Then
                    direction = Trade.TradeExecutionDirection.Sell
                ElseIf _Gap < -0.5 Then
                    direction = Trade.TradeExecutionDirection.Buy
                End If
                Dim entryPrice As Decimal = Decimal.MinValue
                If direction = Trade.TradeExecutionDirection.Buy Then
                    Dim potentialEntryPrice As Decimal = _FractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(potentialEntryPrice, RoundOfType.Floor)
                    entryPrice = potentialEntryPrice + buffer
                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                    Dim potentialEntryPrice As Decimal = _FractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(potentialEntryPrice, RoundOfType.Floor)
                    entryPrice = potentialEntryPrice - buffer
                End If
                If entryPrice <> Decimal.MinValue AndAlso entryPrice <> currentTrade.EntryPrice Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrder(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrder(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function IsSignalCandle(ByVal candle As Payload) As Tuple(Of Boolean, String)
        Dim ret As Tuple(Of Boolean, String) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If IsFractalBreakoutDone(candle.PayloadDate) Then
                ret = New Tuple(Of Boolean, String)(True, "Second fractal breakout")
            ElseIf IsOpposite2FractalFormed(candle.PayloadDate) Then
                ret = New Tuple(Of Boolean, String)(True, "Opposite 2 fractal formed")
            ElseIf IsFabourableFractalFormed(candle.PayloadDate) Then
                ret = New Tuple(Of Boolean, String)(True, "Fabourable Fractal formed")
            End If
        End If
        Return ret
    End Function

    Private _fractalBreakoutDone As Boolean = False
    Private Function IsFractalBreakoutDone(ByVal currentTime As Date) As Boolean
        Dim ret As Boolean = False
        If Not _fractalBreakoutDone Then
            If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
                Dim fractalPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
                If _Gap > 0.5 Then
                    direction = Trade.TradeExecutionDirection.Sell
                    fractalPayload = _FractalLowPayload
                ElseIf _Gap < -0.5 Then
                    direction = Trade.TradeExecutionDirection.Buy
                    fractalPayload = _FractalHighPayload
                End If
                If fractalPayload IsNot Nothing AndAlso fractalPayload.Count > 0 Then
                    Dim currentFractalStartTime As Date = GetStartTimeOfIndicator(currentTime, fractalPayload)
                    For Each runningPayload In _signalPayload.Keys
                        If runningPayload.Date = _tradingDate.Date AndAlso runningPayload <= currentTime Then
                            Dim indicatorStartTime As Date = GetStartTimeOfIndicator(_signalPayload(runningPayload).PreviousCandlePayload.PayloadDate, fractalPayload)
                            If indicatorStartTime.Date = _tradingDate.Date AndAlso indicatorStartTime <> currentFractalStartTime Then
                                Dim indicatorValue As Decimal = fractalPayload(_signalPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                                If direction = Trade.TradeExecutionDirection.Buy Then
                                    Dim entryPrice As Decimal = indicatorValue + _parentStrategy.CalculateBuffer(indicatorValue, RoundOfType.Floor)
                                    If _signalPayload(runningPayload).High > indicatorValue Then
                                        ret = True
                                        _fractalBreakoutDone = True
                                        Exit For
                                    End If
                                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                                    Dim entryPrice As Decimal = indicatorValue - _parentStrategy.CalculateBuffer(indicatorValue, RoundOfType.Floor)
                                    If _signalPayload(runningPayload).Low < indicatorValue Then
                                        ret = True
                                        _fractalBreakoutDone = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Else
            ret = _fractalBreakoutDone
        End If
        Return ret
    End Function

    Private _oppositeFractal As Boolean = False
    Private Function IsOpposite2FractalFormed(ByVal currentTime As Date) As Boolean
        Dim ret As Boolean = False
        If Not _oppositeFractal Then
            If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
                Dim fractalPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
                If _Gap > 0.5 Then
                    direction = Trade.TradeExecutionDirection.Sell
                    fractalPayload = _FractalHighPayload
                ElseIf _Gap < -0.5 Then
                    direction = Trade.TradeExecutionDirection.Buy
                    fractalPayload = _FractalLowPayload
                End If
                If fractalPayload IsNot Nothing AndAlso fractalPayload.Count > 0 Then
                    Dim fractalValue As Decimal = Decimal.MinValue
                    For Each runningPayload In _signalPayload.Keys
                        If runningPayload.Date = _tradingDate.Date AndAlso runningPayload <= currentTime Then
                            Dim indicatorStartTime As Date = GetStartTimeOfIndicator(runningPayload, fractalPayload)
                            If indicatorStartTime.Date = _tradingDate.Date Then
                                Dim indicatorValue As Decimal = fractalPayload(runningPayload)
                                If fractalValue = Decimal.MinValue Then
                                    fractalValue = indicatorValue
                                Else
                                    If direction = Trade.TradeExecutionDirection.Buy Then
                                        If indicatorValue < fractalValue Then
                                            ret = True
                                            _oppositeFractal = True
                                            Exit For
                                        End If
                                    ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                                        If indicatorValue > fractalValue Then
                                            ret = True
                                            _oppositeFractal = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Else
            ret = _oppositeFractal
        End If
        Return ret
    End Function

    Private _fabourableFractal As Boolean = False
    Private Function IsFabourableFractalFormed(ByVal currentTime As Date) As Boolean
        Dim ret As Boolean = False
        If Not _fabourableFractal Then
            If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
                Dim fractalPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim oppositeFractalPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
                If _Gap > 0.5 Then
                    direction = Trade.TradeExecutionDirection.Sell
                    fractalPayload = _FractalLowPayload
                    oppositeFractalPayload = _FractalHighPayload
                ElseIf _Gap < -0.5 Then
                    direction = Trade.TradeExecutionDirection.Buy
                    fractalPayload = _FractalHighPayload
                    oppositeFractalPayload = _FractalLowPayload
                End If
                If fractalPayload IsNot Nothing AndAlso fractalPayload.Count > 0 Then
                    For Each runningPayload In _signalPayload.Keys
                        If runningPayload.Date = _tradingDate.Date AndAlso runningPayload <= currentTime Then
                            Dim indicatorStartTime As Date = GetStartTimeOfIndicator(runningPayload, fractalPayload)
                            If indicatorStartTime.Date = _tradingDate.Date Then
                                Dim indicatorValue As Decimal = fractalPayload(runningPayload)
                                Dim oppositeIndicatorValue As Decimal = GetIndicatorPreviousValue(runningPayload, oppositeFractalPayload)
                                If direction = Trade.TradeExecutionDirection.Buy Then
                                    If oppositeIndicatorValue > indicatorValue Then
                                        ret = True
                                        _fabourableFractal = True
                                        Exit For
                                    End If
                                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                                    If oppositeIndicatorValue < indicatorValue Then
                                        ret = True
                                        _fabourableFractal = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Else
            ret = _fabourableFractal
        End If
        Return ret
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

    Private Function GetIndicatorPreviousValue(ByVal currentTime As Date, ByVal indicatorPayload As Dictionary(Of Date, Decimal)) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        Dim currentIndicatorValue As Decimal = indicatorPayload(currentTime)
        Dim indicatorStartTime As Date = Date.MinValue
        For Each runningIndicator In indicatorPayload.Keys.OrderByDescending(Function(x)
                                                                                 Return x
                                                                             End Function)
            If runningIndicator <= currentTime Then
                If indicatorPayload(runningIndicator) <> currentIndicatorValue Then
                    ret = indicatorPayload(runningIndicator)
                    Exit For
                End If
            End If
        Next
        Return ret
    End Function

    Private Function GetLastDayLastCandleClose() As Decimal
        Dim ret As Decimal = Decimal.MinValue
        Dim lastDayLastCandle As Payload = Nothing
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            Dim previousDayPayload As IEnumerable(Of KeyValuePair(Of Date, Payload)) = _signalPayload.Where(Function(x)
                                                                                                                Return x.Key.Date <> _tradingDate.Date
                                                                                                            End Function)
            If previousDayPayload IsNot Nothing AndAlso previousDayPayload.Count > 0 Then
                lastDayLastCandle = previousDayPayload.LastOrDefault.Value
            End If
        End If
        If lastDayLastCandle IsNot Nothing Then
            ret = lastDayLastCandle.Close
        End If
        Return ret
    End Function
End Class

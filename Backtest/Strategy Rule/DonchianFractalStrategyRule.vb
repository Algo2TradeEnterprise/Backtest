Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class DonchianFractalStrategyRule
    Inherits StrategyRule

    Private _DonchianHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _DonchianLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _DonchianMiddlePayload As Dictionary(Of Date, Decimal) = Nothing
    Private _FractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _FractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _InvalidInstrument As Boolean = False

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller)
        _InvalidInstrument = False
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.DonchianChannel.CalculateDonchianChannel(50, 50, _signalPayload, _DonchianHighPayload, _DonchianLowPayload, _DonchianMiddlePayload)
        Indicator.FractalBands.CalculateFractal(_signalPayload, _FractalHighPayload, _FractalLowPayload)
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
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Not _InvalidInstrument Then
            Dim signalCandle As Payload = Nothing

            Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
            End If
            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy Then
                    If _FractalHighPayload(signalCandle.PayloadDate) <> _DonchianHighPayload(signalCandle.PayloadDate) AndAlso
                        currentTick.Open < _FractalHighPayload(signalCandle.PayloadDate) Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(_FractalHighPayload(signalCandle.PayloadDate), RoundOfType.Floor)
                        Dim entryPrice As Decimal = ConvertFloorCeling(_FractalHighPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                        parameter = New PlaceOrderParameters With {
                            .EntryPrice = entryPrice + buffer,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = _lotSize,
                            .Stoploss = ConvertFloorCeling(_FractalLowPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing) - buffer,
                            .Target = .EntryPrice + 100000,
                            .Buffer = buffer,
                            .SignalCandle = signalCandle,
                            .OrderType = Strategy.TypeOfOrder.Breakout,
                            .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                            .Supporting2 = ConvertFloorCeling(_FractalLowPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                        }
                    End If
                ElseIf signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell Then
                    If _FractalLowPayload(signalCandle.PayloadDate) <> _DonchianLowPayload(signalCandle.PayloadDate) AndAlso
                        currentTick.Open > _FractalLowPayload(signalCandle.PayloadDate) Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(_FractalLowPayload(signalCandle.PayloadDate), RoundOfType.Floor)
                        Dim entryPrice As Decimal = ConvertFloorCeling(_FractalLowPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                        parameter = New PlaceOrderParameters With {
                            .EntryPrice = entryPrice - buffer,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = _lotSize,
                            .Stoploss = ConvertFloorCeling(_FractalHighPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing) + buffer,
                            .Target = .EntryPrice - 100000,
                            .Buffer = buffer,
                            .SignalCandle = signalCandle,
                            .OrderType = Strategy.TypeOfOrder.Breakout,
                            .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                            .Supporting2 = ConvertFloorCeling(_FractalHighPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                        }
                    End If
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            'If Not _parentStrategy.IsTradeActive(currentTick, Trade.TradeType.MIS) Then
            If Not IsPreviousSignalUsed(currentMinuteCandlePayload.PayloadDate) Then
                Me.MaxLossOfThisStock = (parameter.EntryPrice * parameter.Quantity / _parentStrategy.MarginMultiplier) * _parentStrategy.StoplossMultiplier / 100
                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            End If
            'Else
            '    Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TradeType.MIS)
            '    If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.EntryDirection <> parameter.EntryDirection Then
            '        ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            '    End If
            'End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                    Dim entryPrice As Decimal = Decimal.MinValue
                    If signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(_FractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), RoundOfType.Floor)
                        entryPrice = ConvertFloorCeling(_FractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing) + buffer
                    ElseIf signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(_FractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), RoundOfType.Floor)
                        entryPrice = ConvertFloorCeling(_FractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing) - buffer
                    End If
                    If entryPrice <> Decimal.MinValue AndAlso entryPrice <> currentTrade.EntryPrice Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
            End If
        ElseIf currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If signalCandleSatisfied.Item2 <> currentTrade.EntryDirection Then
                    Dim buffer As Decimal = Decimal.MinValue
                    Dim entryPrice As Decimal = Decimal.MinValue
                    Dim validTrade As Boolean = True
                    If signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy Then
                        buffer = _parentStrategy.CalculateBuffer(_FractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), RoundOfType.Floor)
                        entryPrice = ConvertFloorCeling(_FractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing) + buffer
                        If currentTick.High >= entryPrice Then
                            ret = New Tuple(Of Boolean, String)(True, "Reverse Signal")
                        End If
                    ElseIf signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell Then
                        buffer = _parentStrategy.CalculateBuffer(_FractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), RoundOfType.Floor)
                        entryPrice = ConvertFloorCeling(_FractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing) - buffer
                        If currentTick.Low <= entryPrice Then
                            ret = New Tuple(Of Boolean, String)(True, "Reverse Signal")
                        End If
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
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTrade.EntryPrice, RoundOfType.Floor)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If _FractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) > currentTrade.PotentialStopLoss AndAlso
                    _FractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) < currentTrade.EntryPrice - buffer Then
                    triggerPrice = _FractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) - buffer
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If _FractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) < currentTrade.PotentialStopLoss AndAlso
                    _FractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) > currentTrade.EntryPrice + buffer Then
                    triggerPrice = _FractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) + buffer
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Move to fractal: {0}. Time:{1}", triggerPrice, currentTick.PayloadDate))
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
            Dim fractalDonchianState As Tuple(Of Boolean, Trade.TradeExecutionDirection, Date, Decimal) = IsFractalAndDonchianSame(candle)
            If fractalDonchianState IsNot Nothing AndAlso fractalDonchianState.Item1 Then
                If fractalDonchianState.Item2 = Trade.TradeExecutionDirection.Buy Then
                    If IsSignalStillValid(fractalDonchianState.Item3, candle.PayloadDate, Trade.TradeExecutionDirection.Buy, fractalDonchianState.Item4) Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Buy)
                    End If
                ElseIf fractalDonchianState.Item2 = Trade.TradeExecutionDirection.Sell Then
                    If IsSignalStillValid(fractalDonchianState.Item3, candle.PayloadDate, Trade.TradeExecutionDirection.Sell, fractalDonchianState.Item4) Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Sell)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsFractalAndDonchianSame(ByVal candle As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection, Date, Decimal)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection, Date, Decimal) = Nothing
        If candle IsNot Nothing AndAlso _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            For Each runningPayload In _signalPayload.Keys.OrderByDescending(Function(x)
                                                                                 Return x
                                                                             End Function)
                If runningPayload <= candle.PayloadDate Then
                    Dim highMatched As Boolean = False
                    Dim lowMatched As Boolean = False
                    If _FractalHighPayload(runningPayload) = _DonchianHighPayload(runningPayload) Then
                        highMatched = True
                    End If
                    If _FractalLowPayload(runningPayload) = _DonchianLowPayload(runningPayload) Then
                        lowMatched = True
                    End If
                    If highMatched Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Date, Decimal)(True, Trade.TradeExecutionDirection.Sell, runningPayload, _DonchianHighPayload(runningPayload))
                        Exit For
                    ElseIf lowMatched Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Date, Decimal)(True, Trade.TradeExecutionDirection.Buy, runningPayload, _DonchianLowPayload(runningPayload))
                        Exit For
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Private Function IsSignalStillValid(ByVal startTime As Date, ByVal endTime As Date, ByVal direction As Trade.TradeExecutionDirection, ByVal levelPrice As Decimal) As Boolean
        Dim ret As Boolean = False
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            ret = True
            For Each runningPayload In _signalPayload.Keys
                If runningPayload >= startTime AndAlso runningPayload <= endTime Then
                    If direction = Trade.TradeExecutionDirection.Buy Then
                        If _signalPayload(runningPayload).Close < levelPrice Then
                            ret = False
                        End If
                    ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                        If _signalPayload(runningPayload).Close > levelPrice Then
                            ret = False
                        End If
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Private Function IsPreviousSignalUsed(ByVal currentMinute As Date) As Boolean
        Dim ret As Boolean = False
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            For Each runningPayload In _signalPayload.Keys.OrderByDescending(Function(x)
                                                                                 Return x
                                                                             End Function)
                If runningPayload.Date = _tradingDate.Date AndAlso runningPayload < currentMinute Then
                    Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(_signalPayload(runningPayload).PreviousCandlePayload)
                    If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                        Dim buffer As Decimal = Decimal.MinValue
                        Dim entryPrice As Decimal = Decimal.MinValue
                        Dim validTrade As Boolean = True
                        If signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy Then
                            buffer = _parentStrategy.CalculateBuffer(_FractalHighPayload(_signalPayload(runningPayload).PreviousCandlePayload.PayloadDate), RoundOfType.Floor)
                            entryPrice = ConvertFloorCeling(_FractalHighPayload(_signalPayload(runningPayload).PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing) + buffer
                            validTrade = _signalPayload(runningPayload).Open < entryPrice
                        ElseIf signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell Then
                            buffer = _parentStrategy.CalculateBuffer(_FractalLowPayload(_signalPayload(runningPayload).PreviousCandlePayload.PayloadDate), RoundOfType.Floor)
                            entryPrice = ConvertFloorCeling(_FractalLowPayload(_signalPayload(runningPayload).PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing) - buffer
                            validTrade = _signalPayload(runningPayload).Open > entryPrice
                        End If
                        If validTrade AndAlso IsSignalTriggered(entryPrice, signalCandleSatisfied.Item2, runningPayload, runningPayload) Then
                            ret = True
                            _InvalidInstrument = True
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        Return ret
    End Function
End Class

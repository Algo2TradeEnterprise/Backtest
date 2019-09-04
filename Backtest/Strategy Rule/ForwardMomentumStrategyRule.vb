Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class ForwardMomentumStrategyRule
    Inherits StrategyRule

    Private _EMA50Payload As Dictionary(Of Date, Decimal) = Nothing
    Private _EMA100Payload As Dictionary(Of Date, Decimal) = Nothing

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

        Indicator.EMA.CalculateEMA(50, Payload.PayloadFields.Close, _signalPayload, _EMA50Payload)
        Indicator.EMA.CalculateEMA(100, Payload.PayloadFields.Close, _signalPayload, _EMA100Payload)
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
            Dim signalCandle As Payload = Nothing

            Dim signalCandleSatisfied As Boolean = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
            If signalCandleSatisfied Then
                signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
            Else
                Dim completeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Close)
                If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then
                    Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TradeType.MIS)
                    If lastExecutedTrade IsNot Nothing AndAlso ((lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss AndAlso
                        lastExecutedTrade.PLPoint < 0) OrElse (lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress)) Then
                        Dim oppositeSLTrades As List(Of Trade) = completeTrades.FindAll(Function(x)
                                                                                            Return x.ExitCondition = Trade.TradeExitCondition.StopLoss AndAlso
                                                                                            x.SignalCandle.PayloadDate = lastExecutedTrade.SignalCandle.PayloadDate AndAlso
                                                                                            x.EntryDirection <> lastExecutedTrade.EntryDirection
                                                                                        End Function)
                        If oppositeSLTrades Is Nothing OrElse oppositeSLTrades.Count = 0 Then
                            If currentMinuteCandlePayload.PayloadDate < lastExecutedTrade.SignalCandle.PayloadDate.AddMinutes(10) Then
                                signalCandle = lastExecutedTrade.SignalCandle
                            End If
                        End If
                    End If
                End If
            End If
            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim ema50offsetValue As Decimal = GetOffSetIndicatorValue(10, currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, _EMA50Payload)
                Dim ema100offsetValue As Decimal = GetOffSetIndicatorValue(10, currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, _EMA100Payload)

                Dim longActiveTrades As List(Of Trade) = _parentStrategy.GetOpenActiveTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionDirection.Buy)
                If longActiveTrades Is Nothing OrElse longActiveTrades.Count = 0 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    'If signalCandle.High + buffer < Math.Min(ema50offsetValue, ema100offsetValue) OrElse
                    '    signalCandle.High + buffer > Math.Max(ema50offsetValue, ema100offsetValue) Then
                    Dim sl As Decimal = Math.Min(GetOffSetIndicatorValue(10, signalCandle.PayloadDate, _EMA50Payload), GetOffSetIndicatorValue(10, signalCandle.PayloadDate, _EMA100Payload))
                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = signalCandle.High + buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = _lotSize,
                                .Stoploss = ConvertFloorCeling(Math.Max(sl, (.EntryPrice - .EntryPrice * 1 / 100)), _parentStrategy.TickSize, RoundOfType.Floor),
                                .Target = .EntryPrice + ConvertFloorCeling(.EntryPrice * 50 / 100, _parentStrategy.TickSize, RoundOfType.Celing),
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToShortTimeString
                            }
                    'End If
                End If

                Dim shortActiveTrades As List(Of Trade) = _parentStrategy.GetOpenActiveTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionDirection.Sell)
                If shortActiveTrades Is Nothing OrElse shortActiveTrades.Count = 0 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    'If signalCandle.Low - buffer < Math.Min(ema50offsetValue, ema100offsetValue) OrElse
                    '    signalCandle.Low - buffer > Math.Max(ema50offsetValue, ema100offsetValue) Then
                    Dim sl As Decimal = Math.Max(GetOffSetIndicatorValue(10, signalCandle.PayloadDate, _EMA50Payload), GetOffSetIndicatorValue(10, signalCandle.PayloadDate, _EMA100Payload))
                    parameter2 = New PlaceOrderParameters With {
                                .EntryPrice = signalCandle.Low - buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = _lotSize,
                                .Stoploss = ConvertFloorCeling(Math.Min(sl, (.EntryPrice + .EntryPrice * 1 / 100)), _parentStrategy.TickSize, RoundOfType.Celing),
                                .Target = .EntryPrice - ConvertFloorCeling(.EntryPrice * 50 / 100, _parentStrategy.TickSize, RoundOfType.Celing),
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToShortTimeString
                            }
                    'End If
                End If
            End If
        End If
        Dim orderList As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            If orderList Is Nothing Then orderList = New List(Of PlaceOrderParameters)
            orderList.Add(parameter1)
            If Me.MaxProfitOfThisStock = Decimal.MaxValue Then
                Dim capitalRequired As Decimal = parameter1.EntryPrice * parameter1.Quantity / _parentStrategy.MarginMultiplier
                Me.MaxProfitOfThisStock = capitalRequired * 15 / 100
                Me.MaxLossOfThisStock = capitalRequired * 15 / 100
            End If
        End If
        If parameter2 IsNot Nothing Then
            If orderList Is Nothing Then orderList = New List(Of PlaceOrderParameters)
            orderList.Add(parameter2)
            If Me.MaxProfitOfThisStock = Decimal.MaxValue Then
                Dim capitalRequired As Decimal = parameter2.EntryPrice * parameter2.Quantity / _parentStrategy.MarginMultiplier
                Me.MaxProfitOfThisStock = capitalRequired * 15 / 100
                Me.MaxLossOfThisStock = capitalRequired * 15 / 100
            End If
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
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                If currentMinuteCandlePayload.PayloadDate >= signalCandle.PayloadDate.AddMinutes(10) Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                Else
                    Dim signalCandleSatisfied As Boolean = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
                    If signalCandleSatisfied Then
                        Dim entryPrice As Decimal = Decimal.MinValue
                        signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                        If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                            entryPrice = signalCandle.High + buffer
                        ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                            entryPrice = signalCandle.Low - buffer
                        End If
                        If entryPrice <> Decimal.MinValue AndAlso entryPrice <> currentTrade.EntryPrice Then
                            ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                        End If
                    Else
                        Dim completeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentTick, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Close)
                        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentTick, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                        Dim allTrades As List(Of Trade) = New List(Of Trade)
                        If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then allTrades.AddRange(completeTrades)
                        If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then allTrades.AddRange(inprogressTrades)
                        If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                            Dim sameSignalTrades As List(Of Trade) = allTrades.FindAll(Function(x)
                                                                                           Return x.SignalCandle.PayloadDate = signalCandle.PayloadDate
                                                                                       End Function)
                            If sameSignalTrades IsNot Nothing AndAlso sameSignalTrades.Count > 0 Then
                                Dim longTrades As List(Of Trade) = sameSignalTrades.FindAll(Function(x)
                                                                                                Return x.EntryDirection = Trade.TradeExecutionDirection.Buy
                                                                                            End Function)
                                Dim shortTrades As List(Of Trade) = sameSignalTrades.FindAll(Function(x)
                                                                                                 Return x.EntryDirection = Trade.TradeExecutionDirection.Sell
                                                                                             End Function)
                                If longTrades IsNot Nothing AndAlso longTrades.Count > 0 AndAlso
                                    shortTrades IsNot Nothing AndAlso shortTrades.Count > 0 Then
                                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                                End If
                            End If
                        End If

                        ''New logic
                        'If ret Is Nothing Then
                        '    Dim ema50offsetValue As Decimal = GetOffSetIndicatorValue(10, currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, _EMA50Payload)
                        '    Dim ema100offsetValue As Decimal = GetOffSetIndicatorValue(10, currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, _EMA100Payload)
                        '    If currentTrade.EntryPrice > Math.Min(ema50offsetValue, ema100offsetValue) AndAlso
                        '        currentTrade.EntryPrice < Math.Max(ema50offsetValue, ema100offsetValue) Then
                        '        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                        '    End If
                        'End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim slab As Decimal = 1
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim buffer As Decimal = currentTrade.StoplossBuffer
            Dim gain As Decimal = Math.Abs(currentTick.Open - currentTrade.EntryPrice)
            Dim gainPer As Decimal = gain * 100 / currentTrade.EntryPrice
            If gainPer >= slab + 0.5 Then
                Dim movementMul As Decimal = gainPer * 10 / 5
                Dim movementPer As Decimal = Math.Floor(movementMul - (slab * 2)) * 0.5
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    Dim sl As Decimal = currentTrade.EntryPrice + ConvertFloorCeling(currentTrade.EntryPrice * movementPer / 100, _parentStrategy.TickSize, RoundOfType.Celing)
                    If sl > currentTrade.PotentialStopLoss Then
                        triggerPrice = sl
                    End If
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    Dim sl As Decimal = currentTrade.EntryPrice - ConvertFloorCeling(currentTrade.EntryPrice * movementPer / 100, _parentStrategy.TickSize, RoundOfType.Celing)
                    If sl < currentTrade.PotentialStopLoss Then
                        triggerPrice = sl
                    End If
                End If
            Else
                Dim maxGain As Decimal = Math.Abs(currentTrade.EntryPrice - currentTrade.MaximumDrawUp)
                Dim maxGainPer As Decimal = maxGain * 100 / currentTrade.EntryPrice
                If maxGainPer < 1.5 Then
                    If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        Dim sl As Decimal = Math.Min(GetOffSetIndicatorValue(10, currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, _EMA50Payload), GetOffSetIndicatorValue(10, currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, _EMA100Payload))
                        If sl > currentTrade.PotentialStopLoss Then
                            triggerPrice = ConvertFloorCeling(sl, _parentStrategy.TickSize, RoundOfType.Floor)
                        End If
                    ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        Dim sl As Decimal = Math.Max(GetOffSetIndicatorValue(10, currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, _EMA50Payload), GetOffSetIndicatorValue(10, currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, _EMA100Payload))
                        If sl < currentTrade.PotentialStopLoss Then
                            triggerPrice = ConvertFloorCeling(sl, _parentStrategy.TickSize, RoundOfType.Celing)
                        End If
                    End If
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("{0}. Time:{1}. Gain:{2}", triggerPrice, currentTick.PayloadDate, Math.Round(gainPer, 2)))
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function IsSignalCandle(ByVal candle As Payload) As Boolean
        Dim ret As Boolean = False
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim ema50offsetValue As Decimal = GetOffSetIndicatorValue(10, candle.PayloadDate, _EMA50Payload)
            Dim ema100offsetValue As Decimal = GetOffSetIndicatorValue(10, candle.PayloadDate, _EMA100Payload)
            If candle.High > Math.Max(ema50offsetValue, ema100offsetValue) AndAlso
                candle.Low < Math.Min(ema50offsetValue, ema100offsetValue) Then
                If candle.CandleColor = Color.Green AndAlso
                    candle.Close > Math.Max(ema50offsetValue, ema100offsetValue) AndAlso
                    candle.Open < Math.Max(ema50offsetValue, ema100offsetValue) Then
                    ret = True
                ElseIf candle.CandleColor = Color.Red AndAlso
                    candle.Close < Math.Min(ema50offsetValue, ema100offsetValue) AndAlso
                    candle.Open > Math.Min(ema50offsetValue, ema100offsetValue) Then
                    ret = True
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetOffSetIndicatorValue(ByVal offSet As Integer, ByVal currentTime As Date, ByVal indicatorPayload As Dictionary(Of Date, Decimal)) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If indicatorPayload IsNot Nothing AndAlso indicatorPayload.Count > 0 Then
            Dim count As Integer = 0
            For Each runningPayload In indicatorPayload.Keys.OrderByDescending(Function(x)
                                                                                   Return x
                                                                               End Function)
                If runningPayload < currentTime Then
                    count += 1
                    If count >= offSet Then
                        ret = indicatorPayload(runningPayload)
                        Exit For
                    End If
                End If
            Next
        End If
        Return ret
    End Function
End Class

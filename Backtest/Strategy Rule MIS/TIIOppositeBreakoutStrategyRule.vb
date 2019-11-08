Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class TIIOppositeBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public ModifyStoploss As Boolean
    End Class
#End Region

    Private _TIIPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _SignalLinePayload As Dictionary(Of Date, Decimal) = Nothing
    Private _userInputs As StrategyRuleEntities

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.TrendIntensityIndex.CalculateTII(Payload.PayloadFields.Close, 56, 36, _signalPayload, _TIIPayload, _SignalLinePayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > Math.Abs(_parentStrategy.OverAllLossPerDay) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime Then
            Dim signalCandle As Payload = Nothing
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
            Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If lastExecutedTrade Is Nothing Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                Else
                    'Dim lastSignalStartingPoint As Date = GetStartingPointOfTII(lastExecutedTrade.SignalCandle)
                    'Dim currentSignalStartingPoint As Date = GetStartingPointOfTII(currentMinuteCandlePayload.PreviousCandlePayload)
                    'If lastSignalStartingPoint <> currentSignalStartingPoint Then
                    Dim exitCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedTrade.ExitTime, _signalPayload))
                    If currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate >= exitCandlePayload.PayloadDate Then
                        If lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                            signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                        ElseIf Math.Abs(lastExecutedTrade.PLPoint) < lastExecutedTrade.EntryPrice * _userInputs.TargetMultiplier / 100 Then
                            signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                        End If
                    Else
                        _RequiredCandle = Nothing
                    End If
                    'End If
                End If
            End If
            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    Dim targetPoint As Decimal = Decimal.MinValue
                    If lastExecutedTrade IsNot Nothing Then
                        Dim totalLoss As Decimal = Decimal.MinValue
                        Dim activeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Close)
                        If activeTrades IsNot Nothing AndAlso activeTrades.Count > 0 Then
                            totalLoss = activeTrades.Sum(Function(x)
                                                             Return x.PLAfterBrokerage
                                                         End Function)
                        End If
                        Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(Me._tradingSymbol, signalCandle.High + buffer, _lotSize, Math.Abs(totalLoss), Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                        targetPoint = target - signalCandle.High + buffer
                    Else
                        targetPoint = Math.Max((signalCandle.High + buffer) * _userInputs.TargetMultiplier / 100, (signalCandle.CandleRange + 2 * buffer) * 2)
                    End If
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandle.High + buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = _lotSize,
                        .Stoploss = signalCandle.Low - buffer,
                        .Target = .EntryPrice + ConvertFloorCeling(targetPoint, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                        .Supporting2 = _TIIPayload(signalCandle.PayloadDate)
                    }
                ElseIf signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    Dim targetPoint As Decimal = Decimal.MinValue
                    If lastExecutedTrade IsNot Nothing Then
                        Dim totalLoss As Decimal = Decimal.MinValue
                        Dim activeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Close)
                        If activeTrades IsNot Nothing AndAlso activeTrades.Count > 0 Then
                            totalLoss = activeTrades.Sum(Function(x)
                                                             Return x.PLAfterBrokerage
                                                         End Function)
                        End If
                        Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(Me._tradingSymbol, signalCandle.Low - buffer, _lotSize, Math.Abs(totalLoss), Trade.TradeExecutionDirection.Sell, _parentStrategy.StockType)
                        targetPoint = signalCandle.Low - buffer - target
                    Else
                        targetPoint = Math.Max((signalCandle.Low - buffer) * _userInputs.TargetMultiplier / 100, (signalCandle.CandleRange + 2 * buffer) * 2)
                    End If
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandle.Low - buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                        .Quantity = _lotSize,
                        .Stoploss = signalCandle.High + buffer,
                        .Target = .EntryPrice - ConvertFloorCeling(targetPoint, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                        .Supporting2 = _TIIPayload(signalCandle.PayloadDate)
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                    If signalCandle.PayloadDate <> currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
            End If
        ElseIf currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                    If currentTrade.EntryDirection <> signalCandleSatisfied.Item2 Then
                        ret = New Tuple(Of Boolean, String)(True, "Reverse Signal")
                    End If
                End If
                _RequiredCandle = Nothing
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _userInputs.ModifyStoploss Then
            If _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) <= _parentStrategy.NumberOfTradesPerStockPerDay - 1 Then
                Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                    Dim triggerPrice As Decimal = Decimal.MinValue
                    If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        If currentMinuteCandlePayload.PreviousCandlePayload.Low < currentTrade.EntryPrice - currentTrade.EntryBuffer Then
                            If currentMinuteCandlePayload.PreviousCandlePayload.Low - currentTrade.StoplossBuffer > currentTrade.PotentialStopLoss Then
                                triggerPrice = currentMinuteCandlePayload.PreviousCandlePayload.Low - currentTrade.StoplossBuffer
                            End If
                        End If
                    ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        If currentMinuteCandlePayload.PreviousCandlePayload.High > currentTrade.EntryPrice + currentTrade.EntryBuffer Then
                            If currentMinuteCandlePayload.PreviousCandlePayload.High + currentTrade.StoplossBuffer < currentTrade.PotentialStopLoss Then
                                triggerPrice = currentMinuteCandlePayload.PreviousCandlePayload.High + currentTrade.StoplossBuffer
                            End If
                        End If
                    End If
                    If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                        ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("{0}. Time:{1}", triggerPrice, currentTick.PayloadDate))
                    End If
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

    Private _RequiredCandle As Payload = Nothing
    Private Function IsSignalCandle(ByVal candle As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim tii As Decimal = _TIIPayload(candle.PayloadDate)
            If tii > 80 Then
                If candle.CandleColor = Color.Green Then
                    ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Sell)
                    _RequiredCandle = candle
                Else
                    If _RequiredCandle IsNot Nothing AndAlso _RequiredCandle.CandleColor = Color.Green Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Sell)
                    End If
                End If
            ElseIf tii < 20 Then
                If candle.CandleColor = Color.Red Then
                    ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Buy)
                    _RequiredCandle = candle
                Else
                    If _RequiredCandle IsNot Nothing AndAlso _RequiredCandle.CandleColor = Color.Red Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Buy)
                    End If
                End If
            Else
                If _RequiredCandle IsNot Nothing Then
                    If _RequiredCandle.CandleColor = Color.Red Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Buy)
                    End If
                    If _RequiredCandle.CandleColor = Color.Green Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Sell)
                    End If
                End If
            End If
        End If
        'If ret Is Nothing Then _RequiredCandle = Nothing
        Return ret
    End Function

    Private Function GetStartingPointOfTII(ByVal candle As Payload) As Date
        Dim ret As Date = Date.MinValue
        If _TIIPayload IsNot Nothing AndAlso _TIIPayload.Count > 0 Then
            Dim currentTII As Decimal = _TIIPayload(candle.PayloadDate)
            If currentTII > 80 OrElse currentTII < 20 Then
                For Each runningDate In _TIIPayload.Keys.OrderByDescending(Function(x)
                                                                               Return x
                                                                           End Function)
                    If runningDate <= candle.PayloadDate Then
                        If currentTII > 80 Then
                            If _TIIPayload(runningDate) > 80 Then
                                ret = runningDate
                            Else
                                Exit For
                            End If
                        ElseIf currentTII < 20 Then
                            If _TIIPayload(runningDate) < 20 Then
                                ret = runningDate
                            Else
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function
End Class

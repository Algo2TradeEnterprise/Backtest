Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PreviousDayFactorStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public StoplossPercentage As Decimal
        Public FirstTradeTargetPercentage As Decimal
        Public OnwardTradeTargetPercentage As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _previousDayRange As Decimal = Decimal.MinValue

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Dim eodPayload As Dictionary(Of Date, Payload) = _parentStrategy.Cmn.GetRawPayload(Common.DataBaseTable.EOD_Cash, _tradingSymbol, _tradingDate.AddDays(-10), _tradingDate)
        If eodPayload IsNot Nothing AndAlso eodPayload.Count > 0 Then
            Dim previousDayPayload As Payload = eodPayload.LastOrDefault.Value.PreviousCandlePayload
            If previousDayPayload IsNot Nothing Then
                _previousDayRange = previousDayPayload.CandleRange
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim quantity As Integer = Integer.MinValue
            Dim targetPoint As Decimal = Decimal.MinValue
            Dim slPoint As Decimal = Decimal.MinValue
            Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                slPoint = ConvertFloorCeling(signal.Item2 * _userInputs.StoplossPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If lastExecutedOrder Is Nothing Then
                    quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signal.Item2, signal.Item2 - slPoint, -1000, _parentStrategy.StockType)
                    targetPoint = ConvertFloorCeling(signal.Item2 * _userInputs.FirstTradeTargetPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                ElseIf lastExecutedOrder.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                    quantity = lastExecutedOrder.Quantity * 2
                    targetPoint = ConvertFloorCeling(signal.Item2 * _userInputs.OnwardTradeTargetPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                End If
                signalCandle = signal.Item4
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = 0
                    Dim entry As Decimal = signal.Item2
                    Dim stoploss As Decimal = entry - slPoint
                    Dim target As Decimal = entry + targetPoint
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entry,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = GetDayStartFactor().Item2,
                                    .Supporting2 = GetCurrentLow(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                                }
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = 0
                    Dim entry As Decimal = signal.Item2
                    Dim stoploss As Decimal = entry + slPoint
                    Dim target As Decimal = entry - targetPoint
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entry,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = GetDayStartFactor().Item2,
                                    .Supporting2 = GetCurrentHigh(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                                }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                Dim lastEntryTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.LastTradeEntryTime.Hours, _parentStrategy.LastTradeEntryTime.Minutes, _parentStrategy.LastTradeEntryTime.Seconds)
                If currentTick.PayloadDate > lastEntryTime Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
                If ret Is Nothing Then
                    Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                    If signal IsNot Nothing AndAlso signal.Item1 AndAlso
                        (signal.Item3 <> currentTrade.EntryDirection OrElse currentTrade.EntryPrice <> signal.Item2) Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 AndAlso signal.Item3 <> currentTrade.EntryDirection Then
                Dim triggerPrice As Decimal = Decimal.MinValue
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    If signal.Item2 > currentTrade.EntryPrice - ConvertFloorCeling(currentTrade.EntryPrice * _userInputs.StoplossPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor) Then
                        triggerPrice = signal.Item2 + _parentStrategy.TickSize
                    End If
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    If signal.Item2 < currentTrade.EntryPrice + ConvertFloorCeling(currentTrade.EntryPrice * _userInputs.StoplossPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor) Then
                        triggerPrice = signal.Item2 + _parentStrategy.TickSize
                    End If
                End If
                If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, "Opposite Entry force sl")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim factor As Tuple(Of Decimal, String) = GetDayStartFactor()
            If factor IsNot Nothing Then
                Dim buyPrice As Decimal = GetCurrentLow(candle.PayloadDate) + factor.Item1
                Dim sellPrice As Decimal = GetCurrentHigh(candle.PayloadDate) - factor.Item1
                Dim middle As Decimal = (buyPrice + sellPrice) / 2
                Dim range As Decimal = (buyPrice - middle) * 5 / 100
                If currentTick.Open > middle + range Then
                    ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, buyPrice, Trade.TradeExecutionDirection.Buy, candle)
                ElseIf currentTick.Open < middle - range Then
                    ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, sellPrice, Trade.TradeExecutionDirection.Sell, candle)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetDayStartFactor() As Tuple(Of Decimal, String)
        Dim ret As Tuple(Of Decimal, String) = Nothing
        If _previousDayRange <> Decimal.MinValue Then
            Dim timeToCheck As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, 9, 35, 0)
            Dim high As Decimal = _signalPayload.Max(Function(x)
                                                         If x.Key.Date = _tradingDate.Date AndAlso x.Key <= timeToCheck Then
                                                             Return x.Value.High
                                                         Else
                                                             Return Decimal.MinValue
                                                         End If
                                                     End Function)

            Dim low As Decimal = _signalPayload.Min(Function(x)
                                                        If x.Key.Date = _tradingDate.Date AndAlso x.Key <= timeToCheck Then
                                                            Return x.Value.Low
                                                        Else
                                                            Return Decimal.MaxValue
                                                        End If
                                                    End Function)

            Dim range As Decimal = high - low
            Dim factor As Decimal = Decimal.MinValue
            If range <= _previousDayRange * 0.4333 Then
                factor = _previousDayRange * 0.4333
            ElseIf range <= _previousDayRange * 0.7666 Then
                factor = _previousDayRange * 0.7666
            ElseIf range <= _previousDayRange * 1.35 Then
                factor = _previousDayRange * 1.35
            End If
            ret = New Tuple(Of Decimal, String)(ConvertFloorCeling(factor, _parentStrategy.TickSize, RoundOfType.Floor), String.Format("{0},{1},{2}", high, low, factor))
        End If
        Return ret
    End Function

    Private Function GetCurrentHigh(ByVal currentTime As Date) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        ret = _signalPayload.Max(Function(x)
                                     If x.Key.Date = _tradingDate.Date AndAlso x.Key <= currentTime Then
                                         Return x.Value.High
                                     Else
                                         Return Decimal.MinValue
                                     End If
                                 End Function)

        Return ret
    End Function

    Private Function GetCurrentLow(ByVal currentTime As Date) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        ret = _signalPayload.Min(Function(x)
                                     If x.Key.Date = _tradingDate.Date AndAlso x.Key <= currentTime Then
                                         Return x.Value.Low
                                     Else
                                         Return Decimal.MaxValue
                                     End If
                                 End Function)

        Return ret
    End Function
End Class
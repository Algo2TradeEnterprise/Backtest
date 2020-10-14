Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class UntouchVWAPAdaptiveMartingaleStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxProfitPerTrade As Decimal
        Public MaxLossPerTrade As Decimal
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _vwapPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _previousTradingDay As Date
    Private _slPoint As Decimal

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
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.VWAP.CalculateVWAP(_signalPayload, _vwapPayload)
        _previousTradingDay = _parentStrategy.Cmn.GetPreviousTradingDay(Common.DataBaseTable.Intraday_Cash, _tradingSymbol, _tradingDate)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandle, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandle, _parentStrategy.TradeType)
                If lastOrder IsNot Nothing Then
                    Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastOrder.ExitTime, _signalPayload))
                    If currentMinuteCandle.PayloadDate > exitCandle.PayloadDate Then
                        signalCandle = signal.Item3
                    End If
                Else
                    signalCandle = signal.Item3
                    _slPoint = ConvertFloorCeling(GetAverageHighestATR(signalCandle), _parentStrategy.TickSize, RoundOfType.Celing)
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim entryPrice As Decimal = signal.Item2
                Dim stoploss As Decimal = signal.Item2 - _slPoint
                If signal.Item4 = Trade.TradeExecutionDirection.Sell Then stoploss = signal.Item2 + _slPoint

                Dim iteration As Integer = _parentStrategy.StockNumberOfTrades(_tradingDate, _tradingSymbol) + 1
                Dim lossPL As Decimal = _userInputs.MaxLossPerTrade + _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - Math.Abs(entryPrice - stoploss), lossPL, Trade.TypeOfStock.Cash)
                Dim profitPL As Decimal = _userInputs.MaxProfitPerTrade - _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, profitPL, signal.Item4, Trade.TypeOfStock.Cash)

                Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                            .EntryPrice = entryPrice,
                                                            .EntryDirection = signal.Item4,
                                                            .Quantity = quantity,
                                                            .Stoploss = stoploss,
                                                            .Target = target,
                                                            .Buffer = buffer,
                                                            .SignalCandle = signalCandle,
                                                            .OrderType = Trade.TypeOfOrder.SL,
                                                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                                            .Supporting2 = iteration
                                                        }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If signal.Item3.PayloadDate <> currentTrade.SignalCandle.PayloadDate Then
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim signalCandle As Payload = currentCandle.PreviousCandlePayload
            Dim direction As Trade.TradeExecutionDirection = GetDirection(signalCandle)
            If direction = Trade.TradeExecutionDirection.Buy Then
                Dim entryPrice As Decimal = signalCandle.High + _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, signalCandle, direction)
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                Dim entryPrice As Decimal = signalCandle.Low - _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, signalCandle, direction)
            End If
        End If
        Return ret
    End Function

    Private Function GetDirection(ByVal signalCandle As Payload) As Trade.TradeExecutionDirection
        Dim ret As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
        For Each runningPayload In _signalPayload.OrderByDescending(Function(x)
                                                                        Return x.Key
                                                                    End Function)
            If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key <= signalCandle.PayloadDate Then
                Dim vwap As Decimal = _vwapPayload(runningPayload.Key)
                If runningPayload.Value.Low > vwap Then
                    ret = Trade.TradeExecutionDirection.Buy
                    Exit For
                ElseIf runningPayload.Value.High < vwap Then
                    ret = Trade.TradeExecutionDirection.Sell
                    Exit For
                End If
            End If
        Next
        Return ret
    End Function

    Private Function GetAverageHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 AndAlso signalCandle IsNot Nothing Then
            Dim todayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                 If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                                                     Return x.Value
                                                                 Else
                                                                     Return Decimal.MinValue
                                                                 End If
                                                             End Function)

            Dim previousDayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                       If x.Key.Date = _previousTradingDay.Date Then
                                                                           Return x.Value
                                                                       Else
                                                                           Return Decimal.MinValue
                                                                       End If
                                                                   End Function)

            If todayHighestATR <> Decimal.MinValue AndAlso previousDayHighestATR <> Decimal.MinValue Then
                'ret = (todayHighestATR + previousDayHighestATR) / 2
                ret = Math.Max(todayHighestATR, previousDayHighestATR)
            End If
        End If
        Return ret
    End Function
End Class
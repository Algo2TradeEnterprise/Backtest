Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PreviousDayHKTrendStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public MaxLossPerTrade As Decimal
        Public BreakevenMovement As Boolean
    End Class
#End Region

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _slPoint As Decimal = Decimal.MinValue
    Private _targetPoint As Decimal = Decimal.MinValue
    Private _quantity As Integer = Integer.MinValue
    Private _previousDayHighestATR As Decimal = Decimal.MinValue
    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal direction As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        If direction > 0 Then
            _direction = Trade.TradeExecutionDirection.Buy
        ElseIf direction < 0 Then
            _direction = Trade.TradeExecutionDirection.Sell
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
        Indicator.ATR.CalculateATR(14, _hkPayload, _atrPayload)

        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            Dim firstCandle As Payload = Nothing
            For Each runningPayload In _signalPayload
                If runningPayload.Key.Date = _tradingDate.Date Then
                    If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                        firstCandle = runningPayload.Value
                        Exit For
                    End If
                End If
            Next
            If firstCandle IsNot Nothing AndAlso firstCandle.PreviousCandlePayload IsNot Nothing Then
                _previousDayHighestATR = _atrPayload.Max(Function(x)
                                                             If x.Key.Date >= firstCandle.PreviousCandlePayload.PayloadDate.Date AndAlso
                                                             x.Key.Date < firstCandle.PayloadDate.Date Then
                                                                 Return x.Value
                                                             Else
                                                                 Return Decimal.MinValue
                                                             End If
                                                         End Function)
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item3

                    _slPoint = ConvertFloorCeling(GetAverageHighestATR(signalCandle), _parentStrategy.TickSize, RoundOfType.Celing)
                    _quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signal.Item2, signal.Item2 - _slPoint, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                    _targetPoint = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, signal.Item2, _quantity, Math.Abs(_userInputs.MaxLossPerTrade * _userInputs.TargetMultiplier), Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - signal.Item2
                ElseIf lastExecutedOrder IsNot Nothing Then
                    If Val(lastExecutedOrder.Supporting2) <> signal.Item2 Then
                        signalCandle = signal.Item3
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim slPoint As Decimal = _slPoint
                Dim targetPoint As Decimal = _targetPoint
                Dim quantity As Integer = _quantity
                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = currentTick.Open,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item2
                                }
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = currentTick.Open,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item2
                                }
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter IsNot Nothing Then
            parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _userInputs.BreakevenMovement AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim slPoint As Decimal = _slPoint
            Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * 2, _parentStrategy.TickSize, RoundOfType.Floor)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso currentTrade.PotentialStopLoss < currentTrade.EntryPrice Then
                If currentTick.Open >= currentTrade.EntryPrice + targetPoint Then
                    Dim brkevnPnt As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                    triggerPrice = currentTrade.EntryPrice + brkevnPnt
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso currentTrade.PotentialStopLoss > currentTrade.EntryPrice Then
                If currentTick.Open <= currentTrade.EntryPrice - targetPoint Then
                    Dim brkevnPnt As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                    triggerPrice = currentTrade.EntryPrice - brkevnPnt
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso currentTrade.PotentialStopLoss <> triggerPrice Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, _slPoint)
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim atr As Decimal = ConvertFloorCeling(GetAverageHighestATR(currentCandle.PreviousCandlePayload), _parentStrategy.TickSize, RoundOfType.Celing)
            If _direction = Trade.TradeExecutionDirection.Buy Then
                Dim lowestLow As Decimal = GetLowestLow(currentCandle, currentTick)
                If currentTick.Open >= lowestLow + atr Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, lowestLow, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy)
                End If
            ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                Dim highestHigh As Decimal = GetHighestHigh(currentCandle, currentTick)
                If currentTick.Open <= highestHigh - atr Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, highestHigh, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetAverageHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _previousDayHighestATR <> Decimal.MinValue AndAlso _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            Dim todayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                 If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                                                     Return x.Value
                                                                 Else
                                                                     Return Decimal.MinValue
                                                                 End If
                                                             End Function)
            If todayHighestATR <> Decimal.MinValue Then
                ret = (_previousDayHighestATR + todayHighestATR) / 2
            Else
                ret = _previousDayHighestATR
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestHigh(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        Dim candleHigh As Decimal = _signalPayload.Max(Function(x)
                                                           If x.Key.Date = _tradingDate.Date AndAlso x.Key < currentCandle.PayloadDate Then
                                                               Return x.Value.High
                                                           Else
                                                               Return Decimal.MinValue
                                                           End If
                                                       End Function)
        Dim tickHigh As Decimal = currentCandle.Ticks.Max(Function(x)
                                                              If x.PayloadDate <= currentTick.PayloadDate Then
                                                                  Return x.High
                                                              Else
                                                                  Return Decimal.MinValue
                                                              End If
                                                          End Function)
        ret = Math.Max(candleHigh, tickHigh)
        Return ret
    End Function

    Private Function GetLowestLow(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        Dim candleLow As Decimal = _signalPayload.Min(Function(x)
                                                          If x.Key.Date = _tradingDate.Date AndAlso x.Key < currentCandle.PayloadDate Then
                                                              Return x.Value.Low
                                                          Else
                                                              Return Decimal.MaxValue
                                                          End If
                                                      End Function)
        Dim tickLow As Decimal = currentCandle.Ticks.Min(Function(x)
                                                             If x.PayloadDate <= currentTick.PayloadDate Then
                                                                 Return x.Low
                                                             Else
                                                                 Return Decimal.MaxValue
                                                             End If
                                                         End Function)
        ret = Math.Min(candleLow, tickLow)
        Return ret
    End Function
End Class
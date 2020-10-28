Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class FavourableFractalBreakoutStrategyRule
    Inherits StrategyRule

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _previousTradingDay As Date

    Private ReadOnly _previousLoss As Decimal
    Private ReadOnly _previousIteration As Integer
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal previousLoss As Decimal,
                   ByVal previousIteration As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _previousLoss = previousLoss
        _previousIteration = previousIteration
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
        _previousTradingDay = _signalPayload.Where(Function(x)
                                                       Return x.Key.Date = _tradingDate.Date
                                                   End Function).FirstOrDefault.Value.PreviousCandlePayload.PayloadDate
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item3
            End If

            If signalCandle IsNot Nothing Then
                Dim entryPrice As Decimal = signal.Item2
                Dim slPoint As Decimal = ConvertFloorCeling(GetAverageHighestATR(signalCandle) * 1.5, _parentStrategy.TickSize, RoundOfType.Celing)
                Dim lossPL As Decimal = -500 + _previousLoss
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, lossPL, Trade.TypeOfStock.Cash)
                Dim profitPL As Decimal = 500 - _previousLoss
                Dim targetPoint As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, profitPL, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - entryPrice

                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _previousIteration + 1,
                                    .Supporting3 = signal.Item5
                                }
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _previousIteration + 1,
                                    .Supporting3 = signal.Item5
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryDirection <> signal.Item4 OrElse currentTrade.EntryPrice <> signal.Item2 Then
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

    Private _potentialBuy As Decimal = Decimal.MinValue
    Private _potentialSell As Decimal = Decimal.MinValue
    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim signalCandle As Payload = currentCandle.PreviousCandlePayload
            If signalCandle IsNot Nothing AndAlso signalCandle.PreviousCandlePayload IsNot Nothing AndAlso
                signalCandle.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Close, RoundOfType.Floor)
                If _fractalHighPayload(signalCandle.PayloadDate) < _fractalHighPayload(signalCandle.PreviousCandlePayload.PayloadDate) Then
                    _potentialBuy = _fractalHighPayload(signalCandle.PayloadDate) + buffer
                ElseIf _fractalLowPayload(signalCandle.PayloadDate) > _fractalLowPayload(signalCandle.PreviousCandlePayload.PayloadDate) Then
                    _potentialSell = _fractalLowPayload(signalCandle.PayloadDate) - buffer
                End If

                If _potentialBuy <> Decimal.MinValue AndAlso _potentialSell <> Decimal.MinValue Then
                    Dim middle As Decimal = (_potentialBuy + _potentialSell) / 2
                    Dim range As Decimal = _potentialBuy - middle
                    If currentTick.Open >= middle + range * 60 / 100 Then
                        ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, _potentialBuy, currentCandle, Trade.TradeExecutionDirection.Buy, "")
                    ElseIf currentTick.Open <= middle - range * 60 / 100 Then
                        ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, _potentialSell, currentCandle, Trade.TradeExecutionDirection.Sell, "")
                    End If
                ElseIf _potentialBuy <> Decimal.MinValue Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, _potentialBuy, currentCandle, Trade.TradeExecutionDirection.Buy, "")
                ElseIf _potentialSell <> Decimal.MinValue Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, _potentialSell, currentCandle, Trade.TradeExecutionDirection.Sell, "")
                End If
            End If
        End If
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
            If todayHighestATR <> Decimal.MinValue Then
                ret = (todayHighestATR + previousDayHighestATR) / 2
            End If
        End If
        Return ret
    End Function
End Class
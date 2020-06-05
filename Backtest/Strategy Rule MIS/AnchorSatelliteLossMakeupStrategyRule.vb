Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class AnchorSatelliteLossMakeupStrategyRule
    Inherits StrategyRule

    Private _swingHighPayload As Dictionary(Of Date, Decimal)
    Private _swingLowPayload As Dictionary(Of Date, Decimal)
    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _firstCandleOfTheDay As Payload = Nothing
    Private _slPoint As Decimal = Decimal.MinValue
    Private _targetPoint As Decimal = Decimal.MinValue
    Private _quantity As Integer = Integer.MinValue

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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.SwingHighLow.CalculateSwingHighLow(_signalPayload, False, _swingHighPayload, _swingLowPayload)

        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = _tradingDate.Date Then
                If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                    _firstCandleOfTheDay = runningPayload.Value
                    Exit For
                End If
            End If
        Next

        _slPoint = ConvertFloorCeling(GetHighestATR(_firstCandleOfTheDay) * 1.5, _parentStrategy.TickSize, RoundOfType.Celing)
        _quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, _firstCandleOfTheDay.Open, _firstCandleOfTheDay.Open - _slPoint, -500, Trade.TypeOfStock.Cash)
        _targetPoint = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, _firstCandleOfTheDay.Open, _quantity, 500, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - _firstCandleOfTheDay.Open
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signal As Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim entryPrice As Decimal = signal.Item3
                    Dim slPoint As Decimal = _slPoint
                    Dim targetPoint As Decimal = _targetPoint
                    Dim targetRemark As String = "Normal"
                    Dim quantity As Integer = _quantity

                    parameter1 = New PlaceOrderParameters With {
                            .EntryPrice = entryPrice,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice - slPoint,
                            .Target = .EntryPrice + 1000000000000,
                            .Buffer = 0,
                            .SignalCandle = signal.Item2,
                            .OrderType = signal.Item4,
                            .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                            .Supporting2 = _slPoint,
                            .Supporting3 = targetRemark
                        }

                    Dim pl As Decimal = GetStockPotentialPL(currentMinuteCandlePayload, Trade.TradeExecutionDirection.Buy)
                    If pl < 0 Then
                        quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice + targetPoint, Math.Abs(pl), Trade.TypeOfStock.Cash)
                        targetRemark = "SL Makeup"

                        parameter2 = New PlaceOrderParameters With {
                        .EntryPrice = entryPrice,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = quantity,
                        .Stoploss = .EntryPrice - slPoint,
                        .Target = .EntryPrice + targetPoint,
                        .Buffer = 0,
                        .SignalCandle = signal.Item2,
                        .OrderType = signal.Item4,
                        .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                        .Supporting2 = _slPoint,
                        .Supporting3 = targetRemark
                    }
                    End If
                End If
            End If
            If Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signal As Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim entryPrice As Decimal = signal.Item3
                    Dim slPoint As Decimal = _slPoint
                    Dim targetPoint As Decimal = _targetPoint
                    Dim targetRemark As String = "Normal"
                    Dim quantity As Integer = _quantity

                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + slPoint,
                                .Target = .EntryPrice - 1000000000000,
                                .Buffer = 0,
                                .SignalCandle = signal.Item2,
                                .OrderType = signal.Item4,
                                .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = _slPoint,
                                .Supporting3 = targetRemark
                            }

                    Dim pl As Decimal = GetStockPotentialPL(currentMinuteCandlePayload, Trade.TradeExecutionDirection.Sell)
                    If pl < 0 Then
                        quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice + targetPoint, Math.Abs(pl), Trade.TypeOfStock.Cash)
                        targetRemark = "SL Makeup"

                        parameter2 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + slPoint,
                                .Target = .EntryPrice - targetPoint,
                                .Buffer = 0,
                                .SignalCandle = signal.Item2,
                                .OrderType = signal.Item4,
                                .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = _slPoint,
                                .Supporting3 = targetRemark
                            }
                    End If
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter1)
        End If
        If parameter2 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter2)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, currentTrade.EntryDirection)
            If signal IsNot Nothing AndAlso signal.Item1 AndAlso currentTrade.EntryPrice <> signal.Item3 Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)
        Dim ret As Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder) = Nothing
        If candle IsNot Nothing Then
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(candle, Trade.TypeOfTrade.MIS, direction)
            If lastExecutedTrade IsNot Nothing Then
                Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedTrade.ExitTime, _signalPayload))
                If exitCandle.PayloadDate <= candle.PayloadDate Then
                    If direction = Trade.TradeExecutionDirection.Buy Then
                        Dim lowestPoint As Decimal = GetLowestPointOfTheDay(exitCandle, candle)
                        'If currentTick.Open < lowestPoint + _slPoint Then
                        '    ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, lowestPoint + _slPoint, Trade.TypeOfOrder.SL)
                        'Else
                        '    ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, currentTick.Open, Trade.TypeOfOrder.Market)
                        'End If
                        If _swingHighPayload(candle.PayloadDate) > lowestPoint Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                            ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, _swingHighPayload(candle.PayloadDate) + buffer, Trade.TypeOfOrder.SL)
                        End If
                    ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                        Dim highestPoint As Decimal = GetHighestPointOfTheDay(exitCandle, candle)
                        'If currentTick.Open > highestPoint - _slPoint Then
                        '    ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, highestPoint - _slPoint, Trade.TypeOfOrder.SL)
                        'Else
                        '    ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, currentTick.Open, Trade.TypeOfOrder.Market)
                        'End If
                        If _swingLowPayload(candle.PayloadDate) < highestPoint Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                            ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, _swingLowPayload(candle.PayloadDate) - buffer, Trade.TypeOfOrder.SL)
                        End If
                    End If
                End If
            Else
                Dim lastTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(candle, Trade.TypeOfTrade.MIS)
                If lastTrade IsNot Nothing AndAlso candle.PayloadDate <> _firstCandleOfTheDay.PayloadDate Then
                    Dim highestPoint As Decimal = GetHighestPointOfTheDay(_signalPayload(_firstCandleOfTheDay.PayloadDate.AddMinutes(1)), candle)
                    Dim lowestPoint As Decimal = GetLowestPointOfTheDay(_signalPayload(_firstCandleOfTheDay.PayloadDate.AddMinutes(1)), candle)
                    If direction = Trade.TradeExecutionDirection.Buy Then
                        'If currentTick.Open < lowestPoint + _slPoint Then
                        '    ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, lowestPoint + _slPoint, Trade.TypeOfOrder.SL)
                        'Else
                        '    ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, currentTick.Open, Trade.TypeOfOrder.Market)
                        'End If
                        If _swingHighPayload(candle.PayloadDate) > lowestPoint Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                            ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, _swingHighPayload(candle.PayloadDate) + buffer, Trade.TypeOfOrder.SL)
                        End If
                    ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                        'If currentTick.Open > highestPoint - _slPoint Then
                        '    ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, highestPoint - _slPoint, Trade.TypeOfOrder.SL)
                        'Else
                        '    ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, currentTick.Open, Trade.TypeOfOrder.Market)
                        'End If
                        If _swingLowPayload(candle.PayloadDate) < highestPoint Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                            ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, _swingLowPayload(candle.PayloadDate) - buffer, Trade.TypeOfOrder.SL)
                        End If
                    End If
                ElseIf lastTrade Is Nothing Then
                    Dim highestPoint As Decimal = GetHighestPointOfTheDay(_firstCandleOfTheDay, candle)
                    Dim lowestPoint As Decimal = GetLowestPointOfTheDay(_firstCandleOfTheDay, candle)
                    Dim differenceFromHighestPoint As Decimal = highestPoint - currentTick.Open
                    Dim differenceFromLowestPoint As Decimal = currentTick.Open - lowestPoint
                    If differenceFromHighestPoint > _slPoint AndAlso differenceFromLowestPoint > _slPoint Then
                        If differenceFromHighestPoint > differenceFromLowestPoint Then
                            If direction = Trade.TradeExecutionDirection.Sell Then ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, currentTick.Open, Trade.TypeOfOrder.Market)
                        ElseIf differenceFromHighestPoint < differenceFromLowestPoint Then
                            If direction = Trade.TradeExecutionDirection.Buy Then ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, currentTick.Open, Trade.TypeOfOrder.Market)
                        End If
                    ElseIf differenceFromLowestPoint > _slPoint AndAlso direction = Trade.TradeExecutionDirection.Buy Then
                        ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, currentTick.Open, Trade.TypeOfOrder.Market)
                    ElseIf differenceFromHighestPoint > _slPoint AndAlso direction = Trade.TradeExecutionDirection.Sell Then
                        ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, currentTick.Open, Trade.TypeOfOrder.Market)
                    Else
                        If direction = Trade.TradeExecutionDirection.Buy Then
                            'If currentTick.Open < lowestPoint + _slPoint Then
                            '    ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, lowestPoint + _slPoint, Trade.TypeOfOrder.SL)
                            'Else
                            '    ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, currentTick.Open, Trade.TypeOfOrder.Market)
                            'End If
                            If _swingHighPayload(candle.PayloadDate) > lowestPoint Then
                                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                                ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, _swingHighPayload(candle.PayloadDate) + buffer, Trade.TypeOfOrder.SL)
                            End If
                        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                            'If currentTick.Open > highestPoint - _slPoint Then
                            '    ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, highestPoint - _slPoint, Trade.TypeOfOrder.SL)
                            'Else
                            '    ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, currentTick.Open, Trade.TypeOfOrder.Market)
                            'End If
                            If _swingLowPayload(candle.PayloadDate) < highestPoint Then
                                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                                ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, candle, _swingLowPayload(candle.PayloadDate) - buffer, Trade.TypeOfOrder.SL)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            If _firstCandleOfTheDay IsNot Nothing AndAlso _firstCandleOfTheDay.PreviousCandlePayload IsNot Nothing Then
                ret = _atrPayload.Max(Function(x)
                                          If x.Key.Date >= _firstCandleOfTheDay.PreviousCandlePayload.PayloadDate.Date AndAlso
                                            x.Key <= signalCandle.PayloadDate Then
                                              Return x.Value
                                          Else
                                              Return Decimal.MinValue
                                          End If
                                      End Function)
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestPointOfTheDay(ByVal startCandle As Payload, ByVal endCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        ret = _signalPayload.Max(Function(x)
                                     If x.Key >= startCandle.PayloadDate AndAlso x.Key <= endCandle.PayloadDate Then
                                         Return x.Value.High
                                     Else
                                         Return Decimal.MinValue
                                     End If
                                 End Function)
        Return ret
    End Function

    Private Function GetLowestPointOfTheDay(ByVal startCandle As Payload, ByVal endCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        ret = _signalPayload.Min(Function(x)
                                     If x.Key >= startCandle.PayloadDate AndAlso x.Key <= endCandle.PayloadDate Then
                                         Return x.Value.Low
                                     Else
                                         Return Decimal.MaxValue
                                     End If
                                 End Function)
        Return ret
    End Function

    Private Function GetStockPotentialPL(ByVal currentMinutePayload As Payload, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = 0
        Dim completeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
        If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then
            ret += completeTrades.Sum(Function(x)
                                          If x.EntryDirection = direction Then
                                              Return x.PLAfterBrokerage
                                          Else
                                              Return 0
                                          End If
                                      End Function)
        End If
        Return ret
    End Function
End Class
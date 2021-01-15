Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers
Imports Backtest.StrategyHelper

Public Class NearestOptionFibStrategy
    Inherits StrategyRule

    Private _entryLevelWithQuantity As Dictionary(Of Decimal, Integer) = Nothing

    Private ReadOnly _previousDayHigh As Decimal
    Private ReadOnly _previousDayLow As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal previousDayHigh As Decimal,
                   ByVal previousDayLow As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _previousDayHigh = previousDayHigh
        _previousDayLow = previousDayLow
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Dim price As Decimal = 100
        If _tradingSymbol.StartsWith("NIFTY") Then
            price = 500
        ElseIf _tradingSymbol.StartsWith("BANKNIFTY") Then
            price = 1500
        End If

        Dim range As Decimal = _previousDayHigh - _previousDayLow
        Dim multiplier As Decimal = Math.Ceiling(price / _previousDayLow)
        _entryLevelWithQuantity = New Dictionary(Of Decimal, Integer) From {
            {ConvertFloorCeling(_previousDayLow + range * 0 / 100, _parentStrategy.TickSize, RoundOfType.Celing), 2 * multiplier * Me.LotSize},
            {ConvertFloorCeling(_previousDayLow + range * 23 / 100, _parentStrategy.TickSize, RoundOfType.Celing), 3 * multiplier * Me.LotSize},
            {ConvertFloorCeling(_previousDayLow + range * 38 / 100, _parentStrategy.TickSize, RoundOfType.Celing), 2 * multiplier * Me.LotSize},
            {ConvertFloorCeling(_previousDayLow + range * 62 / 100, _parentStrategy.TickSize, RoundOfType.Celing), 3 * multiplier * Me.LotSize},
            {ConvertFloorCeling(_previousDayLow + range * 77 / 100, _parentStrategy.TickSize, RoundOfType.Celing), 2 * multiplier * Me.LotSize},
            {ConvertFloorCeling(_previousDayLow + range * 100 / 100, _parentStrategy.TickSize, RoundOfType.Celing), 3 * multiplier * Me.LotSize}
        }
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signal As Tuple(Of Boolean, Decimal, Payload) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim entryPrice As Decimal = signal.Item2
                Dim quantity As Integer = _entryLevelWithQuantity(entryPrice)
                Dim target As Integer = _entryLevelWithQuantity.Where(Function(x)
                                                                          Return x.Key > entryPrice
                                                                      End Function).OrderBy(Function(y)
                                                                                                Return y.Key
                                                                                            End Function).FirstOrDefault.Key
                Dim stoploss As Decimal = entryPrice - ConvertFloorCeling((target - entryPrice) / 2, _parentStrategy.TickSize, RoundOfType.Floor)

                parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = 0,
                                .SignalCandle = signal.Item3,
                                .OrderType = Trade.TypeOfOrder.SL
                            }
            End If
        End If
        If parameter IsNot Nothing Then ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentMinuteCandle IsNot Nothing Then
                Dim signal As Tuple(Of Boolean, Decimal, Payload) = GetEntrySignal(currentMinuteCandle, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    If signal.Item2 <> currentTrade.EntryPrice AndAlso currentTrade.EntryPrice <> currentTick.Open Then
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Payload) = Nothing
        Dim lastTrade As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS)
        Dim startTime As Date = _tradingDate.Date
        If lastTrade IsNot Nothing Then
            startTime = lastTrade.ExitTime
        End If

        Dim higherLevel As Decimal = _entryLevelWithQuantity.Where(Function(x)
                                                                       Return x.Key > currentTick.Open
                                                                   End Function).OrderBy(Function(y)
                                                                                             Return y.Key
                                                                                         End Function).FirstOrDefault.Key
        Dim lowerLevel As Decimal = _entryLevelWithQuantity.Where(Function(x)
                                                                      Return x.Key < currentTick.Open
                                                                  End Function).OrderBy(Function(y)
                                                                                            Return y.Key
                                                                                        End Function).LastOrDefault.Key
        Dim lowestPoint As Decimal = _signalPayload.Min(Function(x)
                                                            If x.Key >= startTime AndAlso x.Key < currentCandle.PayloadDate Then
                                                                Return x.Value.Low
                                                            Else
                                                                Return Decimal.MaxValue
                                                            End If
                                                        End Function)
        If lowestPoint <> Decimal.MaxValue AndAlso lowestPoint <> Decimal.MinValue AndAlso lowestPoint <> 0 AndAlso
            higherLevel <> Decimal.MaxValue AndAlso higherLevel <> Decimal.MinValue AndAlso higherLevel <> 0 AndAlso
            lowerLevel <> Decimal.MaxValue AndAlso lowerLevel <> Decimal.MinValue AndAlso lowerLevel <> 0 Then
            If lowestPoint <= lowerLevel Then
                ret = New Tuple(Of Boolean, Decimal, Payload)(True, higherLevel, currentCandle)
            End If
        End If
        Return ret
    End Function
End Class
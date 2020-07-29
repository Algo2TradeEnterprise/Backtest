Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class SwingAtDayHLStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public ATRMultiplier As Decimal
        Public TargetMultiplier As Decimal
        Public BreakevenMovement As Boolean
        Public BreakevenTargetMultiplier As Decimal
        Public NumberOfTradeOnEachDirection As Integer
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _swingPayload As Dictionary(Of Date, Indicator.Swing) = Nothing

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
        Indicator.SwingHighLow.CalculateSwingHighLow(_signalPayload, False, _swingPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            GetNumberOfTrades(currentMinuteCandle, Trade.TradeExecutionDirection.Buy) < _userInputs.NumberOfTradeOnEachDirection AndAlso
            GetNumberOfTrades(currentMinuteCandle, Trade.TradeExecutionDirection.Sell) < _userInputs.NumberOfTradeOnEachDirection AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date, Date) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If GetNumberOfTrades(currentMinuteCandle, signal.Item5) < _userInputs.NumberOfTradeOnEachDirection Then
                    signalCandle = signal.Item4
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                If signal.Item5 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = signal.Item2
                    Dim stoploss As Decimal = signal.Item3
                    Dim target As Decimal = entryPrice + ConvertFloorCeling((entryPrice - stoploss) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = signal.Item6.ToString("dd-MM-yyyy HH:mm:ss"),
                                .Supporting3 = signal.Item7.ToString("dd-MM-yyyy HH:mm:ss"),
                                .Supporting4 = ConvertFloorCeling(entryPrice - stoploss, _parentStrategy.TickSize, RoundOfType.Floor)
                            }
                ElseIf signal.Item5 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = signal.Item2
                    Dim stoploss As Decimal = signal.Item3
                    Dim target As Decimal = entryPrice - ConvertFloorCeling((stoploss - entryPrice) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = signal.Item6.ToString("dd-MM-yyyy HH:mm:ss"),
                                .Supporting3 = signal.Item7.ToString("dd-MM-yyyy HH:mm:ss"),
                                .Supporting4 = ConvertFloorCeling(stoploss - entryPrice, _parentStrategy.TickSize, RoundOfType.Floor)
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
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentMinuteCandle.PreviousCandlePayload.Close < currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentMinuteCandle.PreviousCandlePayload.Close > currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
            If ret Is Nothing Then
                Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date, Date) = GetEntrySignal(currentMinuteCandle, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    If currentTrade.EntryDirection <> signal.Item3 OrElse
                    currentTrade.EntryPrice <> signal.Item2 OrElse
                    currentTrade.SignalCandle.PayloadDate <> signal.Item4.PayloadDate Then
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
        If _userInputs.BreakevenMovement AndAlso _userInputs.BreakevenTargetMultiplier <> Decimal.MinValue AndAlso
            currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim slPoint As Decimal = currentTrade.Supporting4
            Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.BreakevenTargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
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
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, Math.Abs(triggerPrice - currentTrade.EntryPrice))
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date, Date)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date, Date) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim candle As Payload = currentCandle.PreviousCandlePayload
            Dim swing As Indicator.Swing = _swingPayload(candle.PayloadDate)
            If swing IsNot Nothing AndAlso swing.SwingHighTime.Date = _tradingDate.Date Then
                Dim lowestCandle As Payload = GetLowestCandleOfTheDay(candle)
                If lowestCandle IsNot Nothing Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(swing.SwingHigh, RoundOfType.Floor)
                    Dim entry As Decimal = swing.SwingHigh + buffer
                    Dim stoploss As Decimal = lowestCandle.Low - buffer
                    If entry - stoploss < _atrPayload(candle.PayloadDate) * _userInputs.ATRMultiplier Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date, Date)(True, entry, stoploss, candle, Trade.TradeExecutionDirection.Buy, swing.SwingHighTime, lowestCandle.PayloadDate)
                    End If
                End If
            End If
            If ret Is Nothing AndAlso swing IsNot Nothing AndAlso swing.SwingLowTime.Date = _tradingDate.Date Then
                Dim highestCandle As Payload = GetHighestCandleOfTheDay(candle)
                If highestCandle IsNot Nothing Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(swing.SwingLow, RoundOfType.Floor)
                    Dim entry As Decimal = swing.SwingLow - buffer
                    Dim stoploss As Decimal = highestCandle.High + buffer
                    If stoploss - entry < _atrPayload(candle.PayloadDate) * _userInputs.ATRMultiplier Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date, Date)(True, entry, stoploss, candle, Trade.TradeExecutionDirection.Sell, swing.SwingLowTime, highestCandle.PayloadDate)
                    End If
                End If
            End If
            If ret Is Nothing Then
                Dim lastExitTrade As Trade = _parentStrategy.GetLastExitTradeOfTheStock(candle, Trade.TypeOfTrade.MIS)
                If lastExitTrade IsNot Nothing AndAlso lastExitTrade.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                    Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExitTrade.ExitTime, _signalPayload))
                    If exitCandle.PayloadDate = currentCandle.PayloadDate Then
                        Dim swingTime As Date = Date.ParseExact(lastExitTrade.Supporting2, "dd-MM-yyyy HH:mm:ss", Nothing)
                        Dim hlCandleTime As Date = Date.ParseExact(lastExitTrade.Supporting3, "dd-MM-yyyy HH:mm:ss", Nothing)
                        Dim swingData As Indicator.Swing = _swingPayload(lastExitTrade.SignalCandle.PayloadDate)
                        Dim hlCandle As Payload = _signalPayload(hlCandleTime)
                        Dim buffer As Decimal = lastExitTrade.EntryBuffer
                        If lastExitTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                            If swingData.SwingHighTime = swingTime Then
                                Dim entry As Decimal = swingData.SwingHigh + buffer
                                Dim stoploss As Decimal = hlCandle.Low - buffer
                                ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date, Date)(True, entry, stoploss, lastExitTrade.SignalCandle, Trade.TradeExecutionDirection.Buy, swingData.SwingHighTime, hlCandle.PayloadDate)
                            End If
                        ElseIf lastExitTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                            If swingData.SwingLowTime = swingTime Then
                                Dim entry As Decimal = swingData.SwingLow - buffer
                                Dim stoploss As Decimal = hlCandle.High + buffer
                                ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date, Date)(True, entry, stoploss, lastExitTrade.SignalCandle, Trade.TradeExecutionDirection.Sell, swingData.SwingLowTime, hlCandle.PayloadDate)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestCandleOfTheDay(ByVal signalCandle As Payload) As Payload
        Dim ret As Payload = Nothing
        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = signalCandle.PayloadDate.Date AndAlso runningPayload.Key <= signalCandle.PayloadDate Then
                If ret Is Nothing Then ret = runningPayload.Value
                If runningPayload.Value.High >= ret.High Then
                    ret = runningPayload.Value
                End If
            End If
        Next
        Return ret
    End Function

    Private Function GetLowestCandleOfTheDay(ByVal signalCandle As Payload) As Payload
        Dim ret As Payload = Nothing
        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = signalCandle.PayloadDate.Date AndAlso runningPayload.Key <= signalCandle.PayloadDate Then
                If ret Is Nothing Then ret = runningPayload.Value
                If runningPayload.Value.Low <= ret.Low Then
                    ret = runningPayload.Value
                End If
            End If
        Next
        Return ret
    End Function

    Private Function GetNumberOfTrades(ByVal candle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Integer
        Dim ret As Integer = 0
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        Dim closeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Close)
        Dim allTrades As List(Of Trade) = New List(Of Trade)
        If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then allTrades.AddRange(inprogressTrades)
        If closeTrades IsNot Nothing AndAlso closeTrades.Count > 0 Then allTrades.AddRange(closeTrades)
        If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
            Dim directionTrades As List(Of Trade) = allTrades.FindAll(Function(x)
                                                                          Return x.EntryDirection = direction
                                                                      End Function)

            If directionTrades IsNot Nothing AndAlso directionTrades.Count > 0 Then
                ret = directionTrades.Count
            End If
        End If
        Return ret
    End Function
End Class
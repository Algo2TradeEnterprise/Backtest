Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class DiamondStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public InitialTargetSlabMultiplier As Decimal
        Public PartialTargetSlabMultiplier As Decimal
    End Class
#End Region

    Private _buyLevels As Dictionary(Of Decimal, Integer)
    Private _sellLevels As Dictionary(Of Decimal, Integer)
    Private _slPoint As Decimal
    Private _quantityMultiplier As Decimal

    Private ReadOnly _slab As Decimal
    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal slab As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        _slab = slab / 2
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Dim firstCandleOfTheDay As Payload = _signalPayload.Where(Function(x)
                                                                      Return x.Key.Date = _tradingDate.Date
                                                                  End Function).FirstOrDefault.Value
        If firstCandleOfTheDay IsNot Nothing Then
            _slPoint = firstCandleOfTheDay.CandleRange
            Dim qtyList As List(Of Integer) = New List(Of Integer) From {1, 2, 4, 8, 16, 8, 4, 2, 1}

            _buyLevels = New Dictionary(Of Decimal, Integer)
            _sellLevels = New Dictionary(Of Decimal, Integer)
            Dim ctr As Integer = 0
            For Each runningQty In qtyList
                _buyLevels.Add(firstCandleOfTheDay.High + _slab * ctr, runningQty)
                _sellLevels.Add(firstCandleOfTheDay.Low - _slab * ctr, runningQty)

                ctr += 1
            Next

            For qtyMul As Integer = 1 To Integer.MaxValue Step 1
                Dim mul As Integer = qtyMul
                Dim totalCapital As Decimal = _buyLevels.Sum(Function(x)
                                                                 Return x.Key * x.Value * mul / _parentStrategy.MarginMultiplier
                                                             End Function)
                If totalCapital >= 100000 Then
                    _quantityMultiplier = mul
                    Exit For
                End If
            Next
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso Me.EligibleToTakeTrade AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = currentTick
            End If

            If signalCandle IsNot Nothing Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = signal.Item2
                    Dim stoploss As Decimal = entryPrice - _slPoint
                    Dim target As Decimal = entryPrice + ConvertFloorCeling(_slab * _userInputs.InitialTargetSlabMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim quantity As Integer = _buyLevels(entryPrice) * _quantityMultiplier

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = entryPrice,
                                .Supporting2 = signal.Item4,
                                .Supporting3 = _slab,
                                .Supporting4 = _slPoint
                            }
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = signal.Item2
                    Dim stoploss As Decimal = entryPrice + _slPoint
                    Dim target As Decimal = entryPrice - ConvertFloorCeling(_slab * _userInputs.InitialTargetSlabMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim quantity As Integer = _sellLevels(entryPrice) * _quantityMultiplier

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = entryPrice,
                                .Supporting2 = signal.Item4,
                                .Supporting3 = _slab,
                                .Supporting4 = _slPoint
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
        If currentTrade IsNot Nothing Then
            If currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
                Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandle, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 AndAlso signal.Item3 <> currentTrade.EntryDirection Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                Else
                    Dim lastExitTrade As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS, currentTrade.EntryDirection)
                    If lastExitTrade IsNot Nothing AndAlso lastExitTrade.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                        ret = New Tuple(Of Boolean, String)(True, "Group SL Hit")
                    End If
                End If
            ElseIf currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    Dim startPrice As Decimal = _buyLevels.FirstOrDefault.Key
                    If currentTick.Open >= startPrice + _slab * _userInputs.PartialTargetSlabMultiplier Then
                        Dim eligibleTrades As List(Of Decimal) = New List(Of Decimal) From {startPrice, startPrice + _slab * 1, startPrice + _slab * 2, startPrice + _slab * 3}
                        If eligibleTrades.Contains(currentTrade.Supporting1) Then
                            ret = New Tuple(Of Boolean, String)(True, "Partial Target")
                        End If
                    End If
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    Dim startPrice As Decimal = _sellLevels.FirstOrDefault.Key
                    If currentTick.Open <= startPrice - _slab * _userInputs.PartialTargetSlabMultiplier Then
                        Dim eligibleTrades As List(Of Decimal) = New List(Of Decimal) From {startPrice, startPrice - _slab * 1, startPrice - _slab * 2, startPrice - _slab * 3}
                        If eligibleTrades.Contains(currentTrade.Supporting1) Then
                            ret = New Tuple(Of Boolean, String)(True, "Partial Target")
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim lastSlab As Decimal = _buyLevels.FirstOrDefault.Key + Math.Floor((currentTick.Open - _buyLevels.FirstOrDefault.Key) / _slab) * _slab
                If lastSlab - _slPoint > currentTrade.PotentialStopLoss Then
                    triggerPrice = lastSlab - _slPoint
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim lastSlab As Decimal = _sellLevels.FirstOrDefault.Key - Math.Floor((_sellLevels.FirstOrDefault.Key - currentTick.Open) / _slab) * _slab
                If lastSlab + _slPoint < currentTrade.PotentialStopLoss Then
                    triggerPrice = lastSlab + _slPoint
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String)
        Dim ret As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            If _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) Then
                Dim runningTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS)
                Dim nextEntryLevel As Decimal = Decimal.MinValue
                If runningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    If Val(runningTrade.Supporting1) < _buyLevels.LastOrDefault.Key Then
                        nextEntryLevel = _buyLevels.Keys.Where(Function(x)
                                                                   Return x > Val(runningTrade.Supporting1)
                                                               End Function).FirstOrDefault
                    End If
                ElseIf runningTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    If Val(runningTrade.Supporting1) > _sellLevels.LastOrDefault.Key Then
                        nextEntryLevel = _sellLevels.Keys.Where(Function(x)
                                                                    Return x < Val(runningTrade.Supporting1)
                                                                End Function).FirstOrDefault
                    End If
                End If
                If nextEntryLevel <> Decimal.MinValue Then
                    ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String)(True, nextEntryLevel, runningTrade.EntryDirection, runningTrade.Supporting2)
                End If
            Else
                Dim buyPrice As Decimal = _buyLevels.FirstOrDefault.Key
                Dim sellPrice As Decimal = _sellLevels.FirstOrDefault.Key
                Dim middle As Decimal = (buyPrice + sellPrice) / 2
                If currentCandle.PreviousCandlePayload.Close < buyPrice AndAlso currentCandle.PreviousCandlePayload.Close > sellPrice Then
                    Dim range As Decimal = buyPrice - middle
                    If currentTick.Open >= middle + range * 70 / 100 Then
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String)(True, buyPrice, Trade.TradeExecutionDirection.Buy, System.Guid.NewGuid().ToString)
                    ElseIf currentTick.Open <= middle - range * 70 / 100 Then
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String)(True, sellPrice, Trade.TradeExecutionDirection.Sell, System.Guid.NewGuid().ToString)
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class
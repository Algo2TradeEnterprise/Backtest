Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public MustInherit Class MathematicalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum TypeOfTarget
        PL = 1
        INR
    End Enum
    Enum TypeOfStoplossMovement
        Slab = 1
        Breakeven
        MaximizeRiskReward
        None
    End Enum
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetType As TypeOfTarget
        Public TargetMultiplier As Decimal
        Public StoplossMovementType As TypeOfStoplossMovement
        Public BreakevenTargetMultiplier As Decimal
    End Class

    Protected Class EntryDetails
        Public EntryPrice As Decimal = Decimal.MinValue
        Public StoplossPrice As Decimal = Decimal.MinValue
        Public TargetPrice As Decimal = Decimal.MinValue
        Public Quantity As Integer = Integer.MinValue
    End Class
#End Region

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
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsTradeOpen(currentMinuteCandle, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeActive(currentMinuteCandle, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 AndAlso signal.Item2 IsNot Nothing Then
                signalCandle = signal.Item3
            End If

            If signalCandle IsNot Nothing Then
                Dim entryPrice As Decimal = signal.Item2.EntryPrice
                Dim stoploss As Decimal = signal.Item2.StoplossPrice
                Dim target As Decimal = signal.Item2.TargetPrice
                Dim quantity As Integer = signal.Item2.Quantity
                Dim slPoint As Decimal = Decimal.MinValue
                Dim trgtPoint As Decimal = Decimal.MinValue

                If _userInputs.TargetType = TypeOfTarget.PL Then
                    If stoploss <> Decimal.MinValue Then
                        slPoint = Math.Abs(entryPrice - stoploss)
                        trgtPoint = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    ElseIf target <> Decimal.MinValue Then
                        trgtPoint = Math.Abs(entryPrice - target)
                        slPoint = ConvertFloorCeling(trgtPoint / _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    End If
                ElseIf _userInputs.TargetType = TypeOfTarget.INR Then
                    If stoploss <> Decimal.MinValue Then
                        slPoint = Math.Abs(entryPrice - stoploss)
                        Dim pl As Decimal = _parentStrategy.CalculatePL(_tradingSymbol, entryPrice, entryPrice - slPoint, quantity, Me.LotSize, _parentStrategy.StockType)
                        target = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(pl * _userInputs.TargetMultiplier), Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                        trgtPoint = target - entryPrice
                    ElseIf target <> Decimal.MinValue Then
                        trgtPoint = Math.Abs(entryPrice - target)
                        Dim pl As Decimal = _parentStrategy.CalculatePL(_tradingSymbol, entryPrice, entryPrice + trgtPoint, quantity, Me.LotSize, _parentStrategy.StockType)
                        stoploss = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(pl / _userInputs.TargetMultiplier) * -1, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                        slPoint = entryPrice - stoploss
                    End If
                End If

                If quantity > 0 AndAlso slPoint <> Decimal.MinValue AndAlso trgtPoint <> Decimal.MinValue Then
                    If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + trgtPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = signal.Item5,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item6,
                                    .Supporting3 = slPoint
                                }
                    ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - trgtPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = signal.Item5,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item6,
                                    .Supporting3 = slPoint
                                }
                    End If
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
            Dim signal As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 AndAlso signal.Item2 IsNot Nothing Then
                If currentTrade.EntryPrice <> signal.Item2.EntryPrice OrElse
                    currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If _userInputs.StoplossMovementType = TypeOfStoplossMovement.Slab Then
                Dim slPoint As Decimal = currentTrade.Supporting3
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    Dim plPoint As Decimal = currentTick.Open - currentTrade.EntryPrice
                    If plPoint > 0 Then
                        Dim triggerPrice As Decimal = Decimal.MinValue
                        Dim multiplier As Integer = Math.Floor(plPoint / slPoint)
                        If multiplier = 1 Then
                            Dim breakevenPoint As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                            triggerPrice = currentTrade.EntryPrice + breakevenPoint
                        ElseIf multiplier > 1 Then
                            Dim movePoint As Decimal = ConvertFloorCeling(slPoint * (multiplier - 1), _parentStrategy.TickSize, RoundOfType.Floor)
                            triggerPrice = currentTrade.EntryPrice + movePoint
                        End If
                        If triggerPrice <> Decimal.MinValue AndAlso triggerPrice > currentTrade.PotentialStopLoss Then
                            ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Move({0}) at {1}", multiplier, currentTick.PayloadDate.ToString("HH:mm:ss")))
                        End If
                    End If
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    Dim plPoint As Decimal = currentTrade.EntryPrice - currentTick.Open
                    If plPoint > 0 Then
                        Dim triggerPrice As Decimal = Decimal.MinValue
                        Dim multiplier As Integer = Math.Floor(plPoint / slPoint)
                        If multiplier = 1 Then
                            Dim breakevenPoint As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                            triggerPrice = currentTrade.EntryPrice - breakevenPoint
                        ElseIf multiplier > 1 Then
                            Dim movePoint As Decimal = ConvertFloorCeling(slPoint * (multiplier - 1), _parentStrategy.TickSize, RoundOfType.Floor)
                            triggerPrice = currentTrade.EntryPrice - movePoint
                        End If
                        If triggerPrice <> Decimal.MinValue AndAlso triggerPrice < currentTrade.PotentialStopLoss Then
                            ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Move({0}) at {1}", multiplier, currentTick.PayloadDate.ToString("HH:mm:ss")))
                        End If
                    End If
                End If
            ElseIf _userInputs.StoplossMovementType = TypeOfStoplossMovement.Breakeven Then
                Dim slPoint As Decimal = currentTrade.Supporting3
                Dim targetPoint As Decimal = slPoint * _userInputs.BreakevenTargetMultiplier
                Dim triggerPrice As Decimal = Decimal.MinValue
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    If currentTick.Open >= currentTrade.EntryPrice + targetPoint Then
                        Dim breakevenPoint As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                        triggerPrice = currentTrade.EntryPrice + breakevenPoint
                    End If
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    If currentTick.Open <= currentTrade.EntryPrice - targetPoint Then
                        Dim breakevenPoint As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                        triggerPrice = currentTrade.EntryPrice - breakevenPoint
                    End If
                End If
                If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Breakeven at {0}", currentTick.PayloadDate.ToString("HH:mm:ss")))
                End If
            ElseIf _userInputs.StoplossMovementType = TypeOfStoplossMovement.MaximizeRiskReward Then
                Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                If currentMinuteCandle.PayloadDate > currentTrade.EntryTime Then
                    Dim slPoint As Decimal = currentTrade.Supporting3
                    Dim triggerPrice As Decimal = Decimal.MinValue
                    If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        If currentMinuteCandle.PreviousCandlePayload.Close > currentTrade.EntryPrice Then
                            triggerPrice = currentTrade.EntryPrice - ConvertFloorCeling(slPoint / 2, _parentStrategy.TickSize, RoundOfType.Celing)
                        End If
                    ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        If currentMinuteCandle.PreviousCandlePayload.Close < currentTrade.EntryPrice Then
                            triggerPrice = currentTrade.EntryPrice + ConvertFloorCeling(slPoint / 2, _parentStrategy.TickSize, RoundOfType.Celing)
                        End If
                    End If
                    If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                        ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Moved at {0}", currentTick.PayloadDate.ToString("HH:mm:ss")))
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

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Protected MustOverride Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)

    'Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)
    '    'Condition,EntryDetails,SignalCandle,Direction,OrderType,Remark
    '    Dim ret As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String) = Nothing
    '    If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then

    '    End If
    '    Return ret
    'End Function
End Class
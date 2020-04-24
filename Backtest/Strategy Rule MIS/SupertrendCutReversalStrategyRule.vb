Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class SupertrendCutReversalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public BreakevenMovement As Boolean
        Public StoplossATRMultiplier As Decimal
        Public TargetATRMultiplier As Decimal
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _supertrendPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _supertrendColorPayload As Dictionary(Of Date, Color) = Nothing

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

        Indicator.Supertrend.CalculateSupertrend(7, 3, _signalPayload, _supertrendPayload, _supertrendColorPayload)
        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
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
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item2
                ElseIf lastExecutedOrder IsNot Nothing Then
                    Dim lastOrderExitTime As Date = _parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedOrder.ExitTime, _signalPayload)
                    If signal.Item2.PayloadDate >= lastOrderExitTime Then
                        signalCandle = signal.Item2
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim stoploss As Decimal = Math.Min(signalCandle.Low, signalCandle.PreviousCandlePayload.Low) - buffer
                    If stoploss < entryPrice Then
                        Dim slPoint As Decimal = entryPrice - stoploss
                        If slPoint <= _atrPayload(signalCandle.PayloadDate) * _userInputs.StoplossATRMultiplier Then
                            Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                            Dim targetRemark As String = "Normal"
                            If targetPoint < ConvertFloorCeling(_atrPayload(signalCandle.PayloadDate) * _userInputs.TargetATRMultiplier, _parentStrategy.TickSize, RoundOfType.Celing) Then
                                targetPoint = ConvertFloorCeling(_atrPayload(signalCandle.PayloadDate) * _userInputs.TargetATRMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                                targetRemark = "ATR"
                            End If
                            Dim target As Decimal = entryPrice + targetPoint
                            Dim quantity As Integer = Me.LotSize

                            parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _atrPayload(signalCandle.PayloadDate),
                                    .Supporting3 = targetRemark
                                }
                        Else
                            Console.WriteLine(String.Format("Trade neglected for sl. Signal candle:{0}", signalCandle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss")))
                        End If
                    End If
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim stoploss As Decimal = Math.Max(signalCandle.High, signalCandle.PreviousCandlePayload.High) + buffer
                    If stoploss > entryPrice Then
                        Dim slPoint As Decimal = stoploss - entryPrice
                        If slPoint <= _atrPayload(signalCandle.PayloadDate) * _userInputs.StoplossATRMultiplier Then
                            Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                            Dim targetRemark As String = "Normal"
                            If targetPoint < ConvertFloorCeling(_atrPayload(signalCandle.PayloadDate) * _userInputs.TargetATRMultiplier, _parentStrategy.TickSize, RoundOfType.Floor) Then
                                targetPoint = ConvertFloorCeling(_atrPayload(signalCandle.PayloadDate) * _userInputs.TargetATRMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                                targetRemark = "ATR"
                            End If
                            Dim target As Decimal = entryPrice - targetPoint
                            Dim quantity As Integer = Me.LotSize

                            parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _atrPayload(signalCandle.PayloadDate),
                                    .Supporting3 = targetRemark
                                }
                        Else
                            Console.WriteLine(String.Format("Trade neglected for sl. Signal candle:{0}", signalCandle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss")))
                        End If
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
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _userInputs.BreakevenMovement AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim slPoint As Decimal = Math.Round(Math.Abs(currentTrade.EntryPrice - currentTrade.PotentialStopLoss), 2)
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentTick.Open >= currentTrade.EntryPrice + slPoint Then
                    triggerPrice = currentTrade.EntryPrice + _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Trade.TradeExecutionDirection.Buy, Me.LotSize, _parentStrategy.StockType)
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentTick.Open <= currentTrade.EntryPrice - slPoint Then
                    triggerPrice = currentTrade.EntryPrice - _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Trade.TradeExecutionDirection.Sell, Me.LotSize, _parentStrategy.StockType)
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("({0})Breakeven at {1}", slPoint, currentTick.PayloadDate.ToString("HH:mm:ss")))
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing AndAlso
            Not candle.DeadCandle AndAlso Not candle.PreviousCandlePayload.DeadCandle Then
            If candle.CandleColor = Color.Green Then
                If _supertrendColorPayload(candle.PayloadDate) = Color.Green AndAlso
                    _supertrendColorPayload(candle.PreviousCandlePayload.PayloadDate) = Color.Green AndAlso
                    candle.PreviousCandlePayload.Low < _supertrendPayload(candle.PreviousCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, candle, Trade.TradeExecutionDirection.Buy)
                End If
            ElseIf candle.CandleColor = Color.Red Then
                If _supertrendColorPayload(candle.PayloadDate) = Color.Red AndAlso
                    _supertrendColorPayload(candle.PreviousCandlePayload.PayloadDate) = Color.Red AndAlso
                    candle.PreviousCandlePayload.High > _supertrendPayload(candle.PreviousCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, candle, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function
End Class
Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class LowRangeCandleBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public BreakevenMovement As Boolean
    End Class
#End Region

    Private _atrUpperBandPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _atrLowerBandPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _atrTrailingStopPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _atrTrailingStopColorPayload As Dictionary(Of Date, Color) = Nothing

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

        Indicator.ATRBands.CalculateATRBands(0.2, 5, Payload.PayloadFields.Close, _signalPayload, _atrUpperBandPayload, _atrLowerBandPayload)
        Indicator.ATRTrailingStop.CalculateATRTrailingStop(14, 5, _signalPayload, _atrTrailingStopPayload, _atrTrailingStopColorPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso Me.EligibleToTakeTrade AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Not _parentStrategy.IsTradeOpen(currentMinuteCandle, Trade.TypeOfTrade.MIS) AndAlso
            Not _parentStrategy.IsTradeActive(currentMinuteCandle, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Decimal, Decimal, Trade.TradeExecutionDirection) = GetSignalForEntry(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastTrade As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandle, Trade.TypeOfTrade.CNC)
                If lastTrade Is Nothing OrElse lastTrade.ExitTime < currentMinuteCandle.PayloadDate Then
                    signalCandle = signal.Item2
                End If
            End If
            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Close, RoundOfType.Celing)
                Dim tgtPnt As Decimal = ConvertFloorCeling(Math.Abs(signal.Item3 - signal.Item4) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                If signal.Item5 = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                            .EntryPrice = signal.Item3,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = Me.LotSize,
                            .Stoploss = signal.Item4,
                            .Target = .EntryPrice + tgtPnt,
                            .Buffer = buffer,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.SL,
                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                            .Supporting2 = signal.Item3 - signal.Item4
                        }
                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                ElseIf signal.Item5 = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                            .EntryPrice = signal.Item3,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = Me.LotSize,
                            .Stoploss = signal.Item4,
                            .Target = .EntryPrice - tgtPnt,
                            .Buffer = buffer,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.SL,
                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                            .Supporting2 = signal.Item4 - signal.Item3
                        }
                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing And currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Payload, Decimal, Decimal, Trade.TradeExecutionDirection) = GetSignalForEntry(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 AndAlso signal.Item2.PayloadDate <> currentTrade.SignalCandle.PayloadDate Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            Else
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                    currentMinuteCandle.PreviousCandlePayload.Close <= currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                    currentMinuteCandle.PreviousCandlePayload.Close >= currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _userInputs.BreakevenMovement AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim slPoint As Decimal = currentTrade.Supporting2
            Dim potentialSL As Decimal = Decimal.MinValue
            Dim brkevenPoint As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                currentTick.Open >= currentTrade.EntryPrice + slPoint Then
                potentialSL = currentTrade.EntryPrice + brkevenPoint
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                currentTick.Open <= currentTrade.EntryPrice - slPoint Then
                potentialSL = currentTrade.EntryPrice - brkevenPoint
            End If
            If potentialSL <> Decimal.MinValue AndAlso potentialSL <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, potentialSL, "Breakeven Movement")
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

    Private Function GetSignalForEntry(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Decimal, Decimal, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Payload, Decimal, Decimal, Trade.TradeExecutionDirection) = Nothing
        Dim bandRange As Decimal = _atrUpperBandPayload(currentCandle.PreviousCandlePayload.PayloadDate) - _atrLowerBandPayload(currentCandle.PreviousCandlePayload.PayloadDate)
        If currentCandle.PreviousCandlePayload.CandleRange >= _parentStrategy.TickSize AndAlso currentCandle.PreviousCandlePayload.Volume > 0 AndAlso
            currentCandle.PreviousCandlePayload.CandleRange <= bandRange Then
            Dim signalCandle As Payload = currentCandle.PreviousCandlePayload
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Close, RoundOfType.Floor)
            Dim buyPrice As Decimal = signalCandle.High + buffer
            Dim sellPrice As Decimal = signalCandle.Low - buffer

            If _atrTrailingStopColorPayload(signalCandle.PayloadDate) = Color.Green Then
                ret = New Tuple(Of Boolean, Payload, Decimal, Decimal, Trade.TradeExecutionDirection)(True, signalCandle, buyPrice, sellPrice, Trade.TradeExecutionDirection.Buy)
            ElseIf _atrTrailingStopColorPayload(signalCandle.PayloadDate) = Color.Red Then
                ret = New Tuple(Of Boolean, Payload, Decimal, Decimal, Trade.TradeExecutionDirection)(True, signalCandle, sellPrice, buyPrice, Trade.TradeExecutionDirection.Sell)
            End If
        End If
        Return ret
    End Function
End Class
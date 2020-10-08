Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class AnchoredVWAPDirectionStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
        Public ATRMultiplier As Decimal
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _anchoredVWAPPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _previousTradingDay As Date

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

        _previousTradingDay = _parentStrategy.Cmn.GetPreviousTradingDay(Common.DataBaseTable.Intraday_Cash, _tradingSymbol, _tradingDate)
        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.AnchoredVWAP.CalculateAnchoredVWAP(New Date(_previousTradingDay.Year, _previousTradingDay.Month, _previousTradingDay.Day, 9, 15, 0), _signalPayload, _anchoredVWAPPayload)
        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandle, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item2
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = 0
                Dim entryPrice As Decimal = currentTick.Open
                Dim slPoint As Decimal = ConvertFloorCeling(_atrPayload(signalCandle.PayloadDate) * _userInputs.ATRMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                Dim targetPoint As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - entryPrice

                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item4,
                                    .Supporting3 = Math.Round(_atrPayload(signalCandle.PayloadDate), 2)
                                }
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item4,
                                    .Supporting3 = Math.Round(_atrPayload(signalCandle.PayloadDate), 2)
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim direction As Trade.TradeExecutionDirection = GetEntryDirection(currentCandle, currentTick)
            If direction = Trade.TradeExecutionDirection.Buy Then
                If currentCandle.PreviousCandlePayload.Close < _fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String)(True, currentCandle.PreviousCandlePayload, direction, "")
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                If currentCandle.PreviousCandlePayload.Close > _fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String)(True, currentCandle.PreviousCandlePayload, direction, "")
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetEntryDirection(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Trade.TradeExecutionDirection
        Dim ret As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
        For Each runningPayload In _signalPayload.OrderByDescending(Function(x)
                                                                        Return x.Key
                                                                    End Function)
            If runningPayload.Key < currentCandle.PayloadDate Then
                If _fractalHighPayload(runningPayload.Key) > _anchoredVWAPPayload(runningPayload.Key) AndAlso
                    _fractalLowPayload(runningPayload.Key) > _anchoredVWAPPayload(runningPayload.Key) Then
                    ret = Trade.TradeExecutionDirection.Buy
                    Exit For
                ElseIf _fractalHighPayload(runningPayload.Key) < _anchoredVWAPPayload(runningPayload.Key) AndAlso
                    _fractalLowPayload(runningPayload.Key) < _anchoredVWAPPayload(runningPayload.Key) Then
                    ret = Trade.TradeExecutionDirection.Sell
                    Exit For
                End If
            End If
        Next
        Return ret
    End Function
End Class
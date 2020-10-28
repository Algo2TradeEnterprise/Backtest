Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class OpeningSlabBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public SlabMultiplier As Decimal
        Public StoplossMultiplier As Decimal
        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _slab As Decimal
    'Private ReadOnly _previousLoss As Decimal
    Private ReadOnly _previousIteration As Integer
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal slab As Decimal,
                   ByVal previousLoss As Decimal,
                   ByVal previousIteration As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        _slab = slab
        '_previousLoss = previousLoss
        _previousIteration = previousIteration
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
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
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item3
            End If

            If signalCandle IsNot Nothing Then
                Dim entryPrice As Decimal = signal.Item2
                Dim slPoint As Decimal = ConvertFloorCeling(_slab * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                Dim lossPL As Decimal = Math.Pow(2, _previousIteration) * 500 * -1
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, lossPL, Trade.TypeOfStock.Cash)
                Dim profitPL As Decimal = Math.Pow(2, _previousIteration) * 500 * _userInputs.TargetMultiplier
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
                                    .Supporting3 = _slab
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
                                    .Supporting3 = _slab
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
            Dim currentCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryDirection <> signal.Item4 Then
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
    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If _potentialBuy = Decimal.MinValue OrElse _potentialSell = Decimal.MinValue Then
            _potentialBuy = currentCandle.Open + _slab * _userInputs.SlabMultiplier
            _potentialSell = currentCandle.Open - _slab * _userInputs.SlabMultiplier
        End If
        Dim middle As Decimal = (_potentialBuy + _potentialSell) / 2
        Dim range As Decimal = _potentialBuy - middle
        If currentTick.Open >= middle + range * 60 / 100 Then
            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, _potentialBuy, currentCandle, Trade.TradeExecutionDirection.Buy)
        ElseIf currentTick.Open <= middle - range * 60 / 100 Then
            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, _potentialSell, currentCandle, Trade.TradeExecutionDirection.Sell)
        End If
        Return ret
    End Function
End Class
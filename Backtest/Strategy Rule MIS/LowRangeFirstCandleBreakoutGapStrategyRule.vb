Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class LowRangeFirstCandleBreakoutGapStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public StoplossPerTrade As Decimal
        Public TargetMultiplier As Decimal
        Public BreakevenMultiplier As Decimal
    End Class
#End Region

    Private _firstCandleOfTheDay As Payload = Nothing
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

        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            For Each runningPayload In _signalPayload.Values
                If runningPayload.PayloadDate.Date = _tradingDate.Date AndAlso runningPayload.PreviousCandlePayload IsNot Nothing AndAlso
                    runningPayload.PayloadDate.Date <> runningPayload.PreviousCandlePayload.PayloadDate.Date Then
                    _firstCandleOfTheDay = runningPayload
                    Exit For
                End If
            Next
        End If
        If _firstCandleOfTheDay Is Nothing Then Me.EligibleToTakeTrade = False
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing AndAlso Me.EligibleToTakeTrade AndAlso
            currentCandle.PayloadDate >= _tradeStartTime AndAlso Not _parentStrategy.IsTradeActive(currentCandle, Trade.TypeOfTrade.MIS) AndAlso
            Not _parentStrategy.IsTradeOpen(currentCandle, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = _firstCandleOfTheDay
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Open, RoundOfType.Floor)
            Dim slPoint As Decimal = signalCandle.CandleRange + 2 * buffer
            Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signalCandle.Open, signalCandle.Open - slPoint, _userInputs.StoplossPerTrade, _parentStrategy.StockType)
            'Dim targetPoint As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, signalCandle.Open, quantity, Math.Abs(_userInputs.StoplossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType) - signalCandle.Open
            Dim targetPoint As Decimal = slPoint * _userInputs.TargetMultiplier
            If quantity > 0 Then
                If _firstCandleOfTheDay.Open < _firstCandleOfTheDay.PreviousCandlePayload.Close Then
                    Dim parameter As New PlaceOrderParameters With {
                                .EntryPrice = signalCandle.High + buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - slPoint,
                                .Target = .EntryPrice + targetPoint,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = targetPoint
                            }
                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                ElseIf _firstCandleOfTheDay.Open > _firstCandleOfTheDay.PreviousCandlePayload.Close Then
                    Dim parameter As New PlaceOrderParameters With {
                            .EntryPrice = signalCandle.Low - buffer,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice + slPoint,
                            .Target = .EntryPrice - targetPoint,
                            .Buffer = 0,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.SL,
                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                            .Supporting2 = targetPoint
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
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso _userInputs.BreakevenMultiplier < 1 Then
            Dim targetPoint As Decimal = currentTrade.Supporting2
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                currentTick.Open >= currentTrade.EntryPrice + targetPoint * _userInputs.BreakevenMultiplier Then
                Dim breakevenPoint As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                triggerPrice = currentTrade.EntryPrice + breakevenPoint
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                currentTick.Open <= currentTrade.EntryPrice - targetPoint * _userInputs.BreakevenMultiplier Then
                Dim breakevenPoint As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                triggerPrice = currentTrade.EntryPrice - breakevenPoint
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, "Breakeven Movement")
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
End Class
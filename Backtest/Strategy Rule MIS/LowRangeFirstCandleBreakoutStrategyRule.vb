Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class LowRangeFirstCandleBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public StoplossPerTrade As Decimal
        Public TargetMultiplier As Decimal
        Public BreakevenMultiplier As Decimal
        Public ReverseSignalEntry As Boolean
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
            currentCandle.PayloadDate >= _tradeStartTime AndAlso Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentCandle, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = _firstCandleOfTheDay
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Open, RoundOfType.Floor)
            Dim slPoint As Decimal = signalCandle.CandleRange + 2 * buffer
            Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signalCandle.Open, signalCandle.Open - slPoint, _userInputs.StoplossPerTrade, _parentStrategy.StockType)
            Dim targetPoint As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, signalCandle.Open, quantity, Math.Abs(_userInputs.StoplossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType) - signalCandle.Open
            If quantity > 0 Then
                If _userInputs.ReverseSignalEntry OrElse Not IsTradeRunningOrComplete(currentCandle, Trade.TradeExecutionDirection.Sell) Then
                    If GetHighestPointOfTheDay(currentCandle) - signalCandle.High + buffer < targetPoint AndAlso
                        Not IsTradeTaken(currentCandle, Trade.TradeExecutionDirection.Buy) AndAlso
                        currentTick.Open < signalCandle.High + buffer Then
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
                                .Supporting2 = slPoint
                            }
                        ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                    End If
                End If
                If _userInputs.ReverseSignalEntry OrElse Not IsTradeRunningOrComplete(currentCandle, Trade.TradeExecutionDirection.Buy) Then
                    If signalCandle.Low - buffer - GetLowestPointOfTheDay(currentCandle) < targetPoint AndAlso
                        Not IsTradeTaken(currentCandle, Trade.TradeExecutionDirection.Sell) AndAlso
                        currentTick.Open > signalCandle.Low - buffer Then
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
                            .Supporting2 = slPoint
                        }
                        ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing And currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            If _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _userInputs.ReverseSignalEntry Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Trade")
            ElseIf _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, Trade.TypeOfTrade.MIS) Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Trade")
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso _userInputs.BreakevenMultiplier < 1 Then
            Dim slPoint As Decimal = currentTrade.Supporting2
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                currentTick.Open >= currentTrade.EntryPrice + slPoint * _userInputs.BreakevenMultiplier Then
                Dim breakevenPoint As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                triggerPrice = currentTrade.EntryPrice + breakevenPoint
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                currentTick.Open <= currentTrade.EntryPrice - slPoint * _userInputs.BreakevenMultiplier Then
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
        If currentTrade IsNot Nothing AndAlso _userInputs.ReverseSignalEntry AndAlso _parentStrategy.IsAnyTradeOfTheStockStoplossReached(currentTick, Trade.TypeOfTrade.MIS) Then
            Dim price As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Math.Abs(_userInputs.StoplossPerTrade), currentTrade.EntryDirection, currentTrade.StockType)
            If price <> Decimal.MinValue AndAlso price <> currentTrade.PotentialTarget Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, price, Math.Abs(price - currentTrade.PotentialStopLoss))
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)

    End Function

    Private Function GetHighestPointOfTheDay(ByVal candle As Payload) As Decimal
        Return _signalPayload.Max(Function(x)
                                      If x.Key.Date = _tradingDate.Date AndAlso
                                        x.Key <= candle.PayloadDate Then
                                          Return x.Value.High
                                      Else
                                          Return Decimal.MinValue
                                      End If
                                  End Function)
    End Function

    Private Function GetLowestPointOfTheDay(ByVal candle As Payload) As Decimal
        Return _signalPayload.Min(Function(x)
                                      If x.Key.Date = _tradingDate.Date AndAlso
                                        x.Key <= candle.PayloadDate Then
                                          Return x.Value.Low
                                      Else
                                          Return Decimal.MaxValue
                                      End If
                                  End Function)
    End Function

    Private Function IsTradeTaken(candle As Payload, direction As Trade.TradeExecutionDirection) As Boolean
        Dim ret As Boolean = False
        Dim openTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Open)
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        Dim closeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Close)
        Dim allTrades As New List(Of Trade)
        If openTrades IsNot Nothing Then allTrades.AddRange(openTrades)
        If inprogressTrades IsNot Nothing Then allTrades.AddRange(inprogressTrades)
        If closeTrades IsNot Nothing Then allTrades.AddRange(closeTrades)
        If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
            Dim available As List(Of Trade) = allTrades.FindAll(Function(x)
                                                                    Return x.EntryDirection = direction
                                                                End Function)
            ret = available IsNot Nothing AndAlso available.Count > 0
        End If
        Return ret
    End Function

    Private Function IsTradeRunningOrComplete(candle As Payload, direction As Trade.TradeExecutionDirection) As Boolean
        Dim ret As Boolean = False
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        Dim closeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Close)
        Dim allTrades As New List(Of Trade)
        If inprogressTrades IsNot Nothing Then allTrades.AddRange(inprogressTrades)
        If closeTrades IsNot Nothing Then allTrades.AddRange(closeTrades)
        If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
            Dim available As List(Of Trade) = allTrades.FindAll(Function(x)
                                                                    Return x.EntryDirection = direction
                                                                End Function)
            ret = available IsNot Nothing AndAlso available.Count > 0
        End If
        Return ret
    End Function
End Class
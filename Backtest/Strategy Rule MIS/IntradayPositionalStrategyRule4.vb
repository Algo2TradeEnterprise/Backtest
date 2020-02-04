Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class IntradayPositionalStrategyRule4
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxStoplossPerTrade As Decimal
        Public MaxProfitPerTrade As Decimal
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
        _userInputs = entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signalCandle As Payload = Nothing
                Dim signalCandleSatisfied As Tuple(Of Boolean, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                    signalCandle = signalCandleSatisfied.Item2
                End If
                If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandle.High + buffer
                    Dim slPrice As Decimal = signalCandle.Low - buffer
                    Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, slPrice, Math.Abs(_userInputs.MaxStoplossPerTrade) * -1, _parentStrategy.StockType)
                    Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, _userInputs.MaxProfitPerTrade, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = slPrice,
                                    .Target = targetPrice,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
                End If
            End If
            If parameter Is Nothing AndAlso Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signalCandle As Payload = Nothing
                Dim signalCandleSatisfied As Tuple(Of Boolean, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                    signalCandle = signalCandleSatisfied.Item2
                End If

                If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandle.Low - buffer
                    Dim slPrice As Decimal = signalCandle.High + buffer
                    Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, slPrice, entryPrice, Math.Abs(_userInputs.MaxStoplossPerTrade) * -1, _parentStrategy.StockType)
                    Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, _userInputs.MaxProfitPerTrade, Trade.TradeExecutionDirection.Sell, _parentStrategy.StockType)

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = slPrice,
                                    .Target = targetPrice,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandleSatisfied As Tuple(Of Boolean, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, currentTrade.EntryDirection)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signalCandleSatisfied.Item2.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
            If ret Is Nothing AndAlso
                _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) >= Me._parentStrategy.NumberOfTradesPerStockPerDay Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Payload)
        Dim ret As Tuple(Of Boolean, Payload) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If candle.High < candle.PreviousCandlePayload.High AndAlso
                candle.Low > candle.PreviousCandlePayload.Low Then
                If IsGapConditionSatisfied(direction, candle.PayloadDate) Then
                    direction = Trade.TradeExecutionDirection.None
                End If
            End If
            If direction = Trade.TradeExecutionDirection.Buy Then
                If candle.High < candle.PreviousCandlePayload.High Then
                    ret = New Tuple(Of Boolean, Payload)(True, candle)
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                If candle.Low > candle.PreviousCandlePayload.Low Then
                    ret = New Tuple(Of Boolean, Payload)(True, candle)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsGapConditionSatisfied(ByVal direction As Trade.TradeExecutionDirection, ByVal candleTime As Date) As Boolean
        Dim ret As Boolean = False
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            Dim potentialCandles As IEnumerable(Of Payload) = _signalPayload.Values.Where(Function(x)
                                                                                              Return x.PayloadDate.Date = _tradingDate.Date AndAlso
                                                                                             x.PayloadDate <= candleTime
                                                                                          End Function)
            If potentialCandles IsNot Nothing AndAlso potentialCandles.Count > 0 Then
                Dim firstCandle As Payload = potentialCandles.OrderBy(Function(x)
                                                                          Return x.PayloadDate
                                                                      End Function).FirstOrDefault
                If firstCandle IsNot Nothing AndAlso firstCandle.PreviousCandlePayload IsNot Nothing Then
                    Dim higestHigh As Decimal = potentialCandles.Max(Function(x)
                                                                         Return x.High
                                                                     End Function)
                    Dim lowestLow As Decimal = potentialCandles.Min(Function(x)
                                                                        Return x.Low
                                                                    End Function)
                    If direction = Trade.TradeExecutionDirection.Buy AndAlso
                        firstCandle.Open > firstCandle.PreviousCandlePayload.Close AndAlso lowestLow > firstCandle.PreviousCandlePayload.Close Then
                        ret = True
                    ElseIf direction = Trade.TradeExecutionDirection.Sell AndAlso
                        firstCandle.Open < firstCandle.PreviousCandlePayload.Close AndAlso higestHigh < firstCandle.PreviousCandlePayload.Close Then
                        ret = True
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class
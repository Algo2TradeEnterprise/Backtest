Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class BothDirectionMultiTradesStrategyRule
    Inherits StrategyRule

    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _slPoint As Decimal = Decimal.MinValue
    Private _quantity As Integer = Integer.MinValue

    Private ReadOnly _maxSingleDirectionTradeCount As Integer = 3
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameterList As List(Of PlaceOrderParameters) = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim buyActiveTrades As Integer = GetActiveTradeCount(currentMinuteCandlePayload, Trade.TradeExecutionDirection.Buy)
            Dim sellActiveTrades As Integer = GetActiveTradeCount(currentMinuteCandlePayload, Trade.TradeExecutionDirection.Sell)
            If buyActiveTrades < _maxSingleDirectionTradeCount Then
                Dim signal As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim entryPrice As Decimal = signal.Item2
                    Dim slPoint As Decimal = GetStoploss(currentMinuteCandlePayload.PreviousCandlePayload)
                    Dim quantity As Integer = GetQuantity(currentMinuteCandlePayload.PreviousCandlePayload)
                    For tradeCtr As Integer = 1 To _maxSingleDirectionTradeCount - buyActiveTrades
                        If tradeCtr > 2 Then slPoint = 10000000
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                   .EntryPrice = entryPrice,
                                                                   .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                   .Quantity = quantity,
                                                                   .Stoploss = .EntryPrice - slPoint,
                                                                   .Target = .EntryPrice + 10000000,
                                                                   .Buffer = 0,
                                                                   .SignalCandle = currentMinuteCandlePayload,
                                                                   .OrderType = Trade.TypeOfOrder.Market,
                                                                   .Supporting1 = tradeCtr
                                                               }

                        If parameterList Is Nothing Then parameterList = New List(Of PlaceOrderParameters)
                        parameterList.Add(parameter)
                    Next
                End If
            End If
            If sellActiveTrades < _maxSingleDirectionTradeCount Then
                Dim signal As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim entryPrice As Decimal = signal.Item2
                    Dim slPoint As Decimal = GetStoploss(currentMinuteCandlePayload.PreviousCandlePayload)
                    Dim quantity As Integer = GetQuantity(currentMinuteCandlePayload.PreviousCandlePayload)
                    For tradeCtr As Integer = 1 To _maxSingleDirectionTradeCount - sellActiveTrades
                        If tradeCtr > 2 Then slPoint = 10000000
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                   .EntryPrice = entryPrice,
                                                                   .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                                                   .Quantity = quantity,
                                                                   .Stoploss = .EntryPrice + slPoint,
                                                                   .Target = .EntryPrice - 10000000,
                                                                   .Buffer = 0,
                                                                   .SignalCandle = currentMinuteCandlePayload,
                                                                   .OrderType = Trade.TypeOfOrder.Market,
                                                                   .Supporting1 = tradeCtr
                                                               }

                        If parameterList Is Nothing Then parameterList = New List(Of PlaceOrderParameters)
                        parameterList.Add(parameter)
                    Next
                End If
            End If
        End If
        If parameterList IsNot Nothing AndAlso parameterList.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameterList)
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Decimal)
        Dim ret As Tuple(Of Boolean, Decimal) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            Dim lastExitTrade As Trade = _parentStrategy.GetLastExitTradeOfTheStock(candle, Trade.TypeOfTrade.MIS, direction)
            If direction = Trade.TradeExecutionDirection.Buy Then
                If lastExitTrade IsNot Nothing Then
                    If currentTick.Open >= lastExitTrade.EntryPrice Then
                        ret = New Tuple(Of Boolean, Decimal)(True, currentTick.Open)
                    End If
                Else
                    ret = New Tuple(Of Boolean, Decimal)(True, currentTick.Open)
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                If lastExitTrade IsNot Nothing Then
                    If currentTick.Open <= lastExitTrade.EntryPrice Then
                        ret = New Tuple(Of Boolean, Decimal)(True, currentTick.Open)
                    End If
                Else
                    ret = New Tuple(Of Boolean, Decimal)(True, currentTick.Open)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetActiveTradeCount(ByVal candle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Integer
        Dim ret As Integer = 0
        Dim openTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Open)
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        Dim allTrades As List(Of Trade) = New List(Of Trade)
        If openTrades IsNot Nothing Then allTrades.AddRange(openTrades)
        If inprogressTrades IsNot Nothing Then allTrades.AddRange(inprogressTrades)
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

    Private Function GetHighestATROfTheDay(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        ret = _atrPayload.Max(Function(x)
                                  If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                      Return x.Value
                                  Else
                                      Return Decimal.MinValue
                                  End If
                              End Function)
        Return ret
    End Function

    Private Function GetQuantity(ByVal signalCandle As Payload) As Integer
        If _quantity = Integer.MinValue Then
            Dim slPoint As Decimal = GetStoploss(signalCandle)
            _quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signalCandle.Open, signalCandle.Open - slPoint, -500, Trade.TypeOfStock.Cash)
            Return _quantity
        Else
            Return _quantity
        End If
    End Function

    Private Function GetStoploss(ByVal signalCandle As Payload) As Decimal
        If _slPoint = Decimal.MinValue Then
            _slPoint = ConvertFloorCeling(GetHighestATROfTheDay(signalCandle) * 1.5, _parentStrategy.TickSize, RoundOfType.Celing)
            Return _slPoint
        Else
            Return _slPoint
        End If
    End Function
End Class
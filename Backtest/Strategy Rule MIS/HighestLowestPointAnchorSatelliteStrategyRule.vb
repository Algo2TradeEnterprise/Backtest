Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HighestLowestPointAnchorSatelliteStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public ATRMultiplier As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _firstCandleOfTheDay As Payload = Nothing

    Private _slPoint As Decimal = Decimal.MinValue
    'Private _targetPoint As Decimal = Decimal.MinValue
    Private _quantity As Integer = Integer.MinValue

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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)

        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = _tradingDate.Date Then
                If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                    _firstCandleOfTheDay = runningPayload.Value
                    Exit For
                End If
            End If
        Next

        _slPoint = ConvertFloorCeling(GetHighestATR(_firstCandleOfTheDay) * _userInputs.ATRMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
        _quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, _firstCandleOfTheDay.Open, _firstCandleOfTheDay.Open - _slPoint, -500, Trade.TypeOfStock.Cash)
        '_targetPoint = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, _firstCandleOfTheDay.Open, _quantity, 500, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - _firstCandleOfTheDay.Open
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signal As Tuple(Of Boolean, Payload, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim tradeTag As String = System.Guid.NewGuid.ToString()
                    Dim entryPrice As Decimal = signal.Item3
                    Dim slPoint As Decimal = _slPoint
                    Dim quantity As Integer = _quantity

                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - slPoint,
                                .Target = .EntryPrice + 1000000000000,
                                .Buffer = 0,
                                .SignalCandle = signal.Item2,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = _slPoint,
                                .Supporting3 = "Anchor",
                                .Supporting4 = tradeTag
                            }

                    If GetActiveTradeCount(currentMinuteCandlePayload, Trade.TradeExecutionDirection.Sell) = 1 Then
                        Dim sellActiveTrade As Trade = GetActiveAnchorTrade(currentMinuteCandlePayload, Trade.TradeExecutionDirection.Sell)
                        If sellActiveTrade IsNot Nothing AndAlso sellActiveTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            If entryPrice < sellActiveTrade.EntryPrice Then
                                parameter2 = New PlaceOrderParameters With {
                                            .EntryPrice = entryPrice,
                                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                            .Quantity = quantity,
                                            .Stoploss = .EntryPrice + slPoint,
                                            .Target = .EntryPrice - 1000000000000,
                                            .Buffer = 0,
                                            .SignalCandle = signal.Item2,
                                            .OrderType = Trade.TypeOfOrder.Market,
                                            .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                            .Supporting2 = _slPoint,
                                            .Supporting3 = "Satellite",
                                            .Supporting4 = tradeTag
                                        }
                            End If
                        End If
                    End If
                End If
            End If
            If Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signal As Tuple(Of Boolean, Payload, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim tradeTag As String = System.Guid.NewGuid.ToString()
                    Dim entryPrice As Decimal = signal.Item3
                    Dim slPoint As Decimal = _slPoint
                    Dim quantity As Integer = _quantity

                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + slPoint,
                                .Target = .EntryPrice - 1000000000000,
                                .Buffer = 0,
                                .SignalCandle = signal.Item2,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = _slPoint,
                                .Supporting3 = "Anchor",
                                .Supporting4 = tradeTag
                            }

                    If GetActiveTradeCount(currentMinuteCandlePayload, Trade.TradeExecutionDirection.Buy) = 1 Then
                        Dim buyActiveTrade As Trade = GetActiveAnchorTrade(currentMinuteCandlePayload, Trade.TradeExecutionDirection.Buy)
                        If buyActiveTrade IsNot Nothing AndAlso buyActiveTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            If entryPrice > buyActiveTrade.EntryPrice Then
                                parameter2 = New PlaceOrderParameters With {
                                            .EntryPrice = entryPrice,
                                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                            .Quantity = quantity,
                                            .Stoploss = .EntryPrice - slPoint,
                                            .Target = .EntryPrice + 1000000000000,
                                            .Buffer = 0,
                                            .SignalCandle = signal.Item2,
                                            .OrderType = Trade.TypeOfOrder.Market,
                                            .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                            .Supporting2 = _slPoint,
                                            .Supporting3 = "Satellite",
                                            .Supporting4 = tradeTag
                                        }
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter1)
        End If
        If parameter2 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter2)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso currentTrade.Supporting3 = "Satellite" Then
            Dim direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                direction = Trade.TradeExecutionDirection.Sell
            Else
                direction = Trade.TradeExecutionDirection.Buy
            End If
            Dim similarTrade As Trade = GetOppositeSimilarTagTrade(currentTick, direction, currentTrade.Supporting4)
            If similarTrade IsNot Nothing AndAlso similarTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close AndAlso similarTrade.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                ret = New Tuple(Of Boolean, String)(True, "Opposite Anchor Trade hits sl")
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Payload, Decimal)
        Dim ret As Tuple(Of Boolean, Payload, Decimal) = Nothing
        If candle IsNot Nothing Then
            If direction = Trade.TradeExecutionDirection.Buy Then
                Dim lowestPoint As Decimal = GetLowestPointOfTheDay(currentTick.PayloadDate)
                If lowestPoint <> Decimal.MaxValue AndAlso currentTick.Open >= lowestPoint + _slPoint Then
                    ret = New Tuple(Of Boolean, Payload, Decimal)(True, candle, currentTick.Open)
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                Dim highestPoint As Decimal = GetHighestPointOfTheDay(currentTick.PayloadDate)
                If highestPoint <> Decimal.MinValue AndAlso currentTick.Open <= highestPoint - _slPoint Then
                    ret = New Tuple(Of Boolean, Payload, Decimal)(True, candle, currentTick.Open)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            If _firstCandleOfTheDay IsNot Nothing AndAlso _firstCandleOfTheDay.PreviousCandlePayload IsNot Nothing Then
                ret = _atrPayload.Max(Function(x)
                                          If x.Key.Date >= _firstCandleOfTheDay.PreviousCandlePayload.PayloadDate.Date AndAlso
                                            x.Key < signalCandle.PayloadDate Then
                                              Return x.Value
                                          Else
                                              Return Decimal.MinValue
                                          End If
                                      End Function)
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestPointOfTheDay(ByVal currentTime As Date) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key < currentTime Then
                If runningPayload.Value.Ticks IsNot Nothing AndAlso runningPayload.Value.Ticks.Count > 0 Then
                    For Each runningTick In runningPayload.Value.Ticks
                        If runningTick.PayloadDate < currentTime Then
                            ret = Math.Max(ret, runningTick.Open)
                        End If
                    Next
                End If
            End If
        Next
        Return ret
    End Function

    Private Function GetLowestPointOfTheDay(ByVal currentTime As Date) As Decimal
        Dim ret As Decimal = Decimal.MaxValue
        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key < currentTime Then
                If runningPayload.Value.Ticks IsNot Nothing AndAlso runningPayload.Value.Ticks.Count > 0 Then
                    For Each runningTick In runningPayload.Value.Ticks
                        If runningTick.PayloadDate < currentTime Then
                            ret = Math.Min(ret, runningTick.Open)
                        End If
                    Next
                End If
            End If
        Next
        Return ret
    End Function

    Private Function GetActiveAnchorTrade(ByVal currentMinutePayload As Payload, ByVal direction As Trade.TradeExecutionDirection) As Trade
        Dim ret As Trade = Nothing
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Inprogress)
        If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then
            Dim directionTrades As List(Of Trade) = inprogressTrades.FindAll(Function(x)
                                                                                 Return x.EntryDirection = direction AndAlso x.Supporting3 = "Anchor"
                                                                             End Function)
            If directionTrades IsNot Nothing AndAlso directionTrades.Count > 0 Then
                ret = directionTrades.OrderByDescending(Function(x)
                                                            Return x.EntryTime
                                                        End Function).FirstOrDefault
            End If
        End If
        Return ret
    End Function

    Private Function GetActiveTradeCount(ByVal currentMinutePayload As Payload, ByVal direction As Trade.TradeExecutionDirection) As Integer
        Dim ret As Integer = 0
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Inprogress)
        If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then
            Dim directionTrades As List(Of Trade) = inprogressTrades.FindAll(Function(x)
                                                                                 Return x.EntryDirection = direction
                                                                             End Function)
            If directionTrades IsNot Nothing AndAlso directionTrades.Count > 0 Then
                ret = directionTrades.Count
            End If
        End If
        Return ret
    End Function

    Private Function GetOppositeSimilarTagTrade(ByVal currentMinutePayload As Payload, ByVal direction As Trade.TradeExecutionDirection, ByVal tag As String) As Trade
        Dim ret As Trade = Nothing
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
        Dim closeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
        Dim allTrdaes As List(Of Trade) = New List(Of Trade)
        If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then allTrdaes.AddRange(inprogressTrades)
        If closeTrades IsNot Nothing AndAlso closeTrades.Count > 0 Then allTrdaes.AddRange(closeTrades)
        If allTrdaes IsNot Nothing AndAlso allTrdaes.Count > 0 Then
            ret = allTrdaes.Find(Function(x)
                                     Return x.EntryDirection = direction AndAlso x.Supporting4 = tag AndAlso x.Supporting3 = "Anchor"
                                 End Function)
        End If
        Return ret
    End Function
End Class
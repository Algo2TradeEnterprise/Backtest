Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class MartingaleStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public StoplossATRMultiplier As Decimal
        Public MinimumStoplossPoint As Decimal
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _firstCandleATR As Decimal = Decimal.MinValue
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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)

        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            For Each runningPayload In _signalPayload
                If runningPayload.Key.Date = _tradingDate.Date Then
                    If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                        _firstCandleATR = _atrPayload(runningPayload.Key)
                        Exit For
                    End If
                End If
            Next
        End If
        If _firstCandleATR = Decimal.MinValue Then Throw New ApplicationException("Unable to fetch first candle ATR")
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            GetQuantity(currentMinuteCandlePayload) <> Me.LotSize Then

            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
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
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim slPoint As Decimal = ConvertFloorCeling(_firstCandleATR * _userInputs.StoplossATRMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                    If slPoint <= _userInputs.MinimumStoplossPoint Then
                        slPoint = _userInputs.MinimumStoplossPoint
                    End If
                    Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim targetRemark As String = "Normal"
                    Dim quantity As Integer = Me.LotSize

                    parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = entryPrice - slPoint,
                                    .Target = entryPrice + targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _firstCandleATR,
                                    .Supporting3 = targetRemark
                                }

                    quantity = GetQuantity(currentMinuteCandlePayload)
                    If quantity < 0 Then
                        Dim pl As Decimal = _parentStrategy.StockPLAfterBrokerage(currentMinuteCandlePayload.PayloadDate, currentMinuteCandlePayload.TradingSymbol)
                        Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, Math.Abs(quantity), Math.Abs(pl), Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                        parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = Math.Abs(quantity),
                                    .Stoploss = entryPrice - slPoint,
                                    .Target = targetPrice,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _firstCandleATR,
                                    .Supporting3 = "SL Makeup"
                                }
                    End If
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim slPoint As Decimal = ConvertFloorCeling(_firstCandleATR * _userInputs.StoplossATRMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                    If slPoint <= _userInputs.MinimumStoplossPoint Then
                        slPoint = _userInputs.MinimumStoplossPoint
                    End If
                    Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim targetRemark As String = "Normal"
                    Dim quantity As Integer = Me.LotSize

                    parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = entryPrice + slPoint,
                                    .Target = entryPrice - targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _firstCandleATR,
                                    .Supporting3 = targetRemark
                                }

                    quantity = GetQuantity(currentMinuteCandlePayload)
                    If quantity < 0 Then
                        Dim pl As Decimal = _parentStrategy.StockPLAfterBrokerage(currentMinuteCandlePayload.PayloadDate, currentMinuteCandlePayload.TradingSymbol)
                        Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, Math.Abs(quantity), Math.Abs(pl), Trade.TradeExecutionDirection.Sell, _parentStrategy.StockType)
                        parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = Math.Abs(quantity),
                                    .Stoploss = entryPrice + slPoint,
                                    .Target = targetPrice,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _firstCandleATR,
                                    .Supporting3 = "SL Makeup"
                                }
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            If candle.CandleColor = Color.Green Then
                If candle.CandleWicks.Top <= candle.CandleRange * 20 / 100 Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, candle, Trade.TradeExecutionDirection.Buy)
                End If
            ElseIf candle.CandleColor = Color.Red Then
                If candle.CandleWicks.Bottom <= candle.CandleRange * 20 / 100 Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, candle, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetQuantity(ByVal currentMinuteCandle As Payload) As Integer
        Dim ret As Integer = 0
        Dim closeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinuteCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Close)
        If closeTrades IsNot Nothing AndAlso closeTrades.Count > 0 Then
            For Each runningTrade In closeTrades
                If runningTrade.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                    ret -= runningTrade.Quantity
                ElseIf runningTrade.ExitCondition = Trade.TradeExitCondition.Target Then
                    ret += runningTrade.Quantity
                End If
            Next
        End If
        Return ret
    End Function
End Class
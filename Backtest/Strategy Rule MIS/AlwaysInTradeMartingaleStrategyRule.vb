﻿Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class AlwaysInTradeMartingaleStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private _oneSlabProfit As Decimal

    Private ReadOnly _slab As Decimal
    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal slab As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        _slab = slab
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        Dim parameter3 As PlaceOrderParameters = Nothing
        Dim parameter4 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If Me.MaxProfitOfThisStock = Decimal.MaxValue Then
                Dim price As Decimal = currentTick.Open
                Dim targetPrice As Decimal = price + _slab * _userInputs.TargetMultiplier
                Dim pl As Decimal = _parentStrategy.CalculatePL(_tradingSymbol, price, targetPrice, Me.LotSize, Me.LotSize, _parentStrategy.StockType)
                Me.MaxProfitOfThisStock = pl
                _oneSlabProfit = _parentStrategy.CalculatePL(_tradingSymbol, price, price + _slab, Me.LotSize, Me.LotSize, _parentStrategy.StockType)
                If _oneSlabProfit <= 0 Then
                    Throw New ApplicationException(String.Format("One Slab Profit PL:{0}", _oneSlabProfit))
                End If
            End If
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signal As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2 + buffer
                    Dim slPoint As Decimal = _slab + 2 * buffer
                    Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim targetRemark As String = "Normal"
                    Dim quantity As Integer = Me.LotSize

                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = entryPrice - slPoint,
                                .Target = entryPrice + targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = currentMinuteCandlePayload,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = currentTick.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = targetRemark
                            }

                    Dim pl As Decimal = _parentStrategy.StockPLAfterBrokerage(currentMinuteCandlePayload.PayloadDate, currentMinuteCandlePayload.TradingSymbol)
                    If pl < 0 Then
                        quantity = Math.Ceiling(Math.Abs(pl) / (_oneSlabProfit * 1.5)) * Me.LotSize
                        Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(pl), Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                        parameter3 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = Math.Abs(quantity),
                                    .Stoploss = entryPrice - slPoint,
                                    .Target = targetPrice,
                                    .Buffer = 0,
                                    .SignalCandle = currentMinuteCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = currentTick.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = "SL Makeup"
                                }
                    End If
                End If
            End If
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signal As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2 - buffer
                    Dim slPoint As Decimal = _slab + 2 * buffer
                    Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim targetRemark As String = "Normal"
                    Dim quantity As Integer = Me.LotSize

                    parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = entryPrice + slPoint,
                                    .Target = entryPrice - targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = currentMinuteCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = currentTick.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = targetRemark
                                }

                    Dim pl As Decimal = _parentStrategy.StockPLAfterBrokerage(currentMinuteCandlePayload.PayloadDate, currentMinuteCandlePayload.TradingSymbol)
                    If pl < 0 Then
                        quantity = Math.Ceiling(Math.Abs(pl) / (_oneSlabProfit * 1.5)) * Me.LotSize
                        Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(pl), Trade.TradeExecutionDirection.Sell, _parentStrategy.StockType)
                        parameter4 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = Math.Abs(quantity),
                                    .Stoploss = entryPrice + slPoint,
                                    .Target = targetPrice,
                                    .Buffer = 0,
                                    .SignalCandle = currentMinuteCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = currentTick.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = "SL Makeup"
                                }
                    End If
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter1)
        End If
        If parameter2 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter2)
        End If
        If parameter3 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter3)
        End If
        If parameter4 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter4)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastSpecificTrades(currentTick, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Inprogress)
            If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                If lastExecutedTrade.EntryDirection <> currentTrade.EntryDirection Then
                    If currentTrade.EntryPrice <> lastExecutedTrade.PotentialStopLoss OrElse currentTrade.EntryTime < lastExecutedTrade.EntryTime Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim slabLevel As Decimal = GetSlabBasedLevel(currentTick.Open, Trade.TradeExecutionDirection.Sell)
                If slabLevel - _slab - currentTrade.StoplossBuffer > currentTrade.PotentialStopLoss Then
                    triggerPrice = slabLevel - _slab - currentTrade.StoplossBuffer
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim slabLevel As Decimal = GetSlabBasedLevel(currentTick.Open, Trade.TradeExecutionDirection.Buy)
                If slabLevel + _slab + currentTrade.StoplossBuffer < currentTrade.PotentialStopLoss Then
                    triggerPrice = slabLevel + _slab + currentTrade.StoplossBuffer
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Moved at {0}", triggerPrice))
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

    Private Function GetSlabBasedLevel(ByVal price As Decimal, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If direction = Trade.TradeExecutionDirection.Buy Then
            ret = Math.Ceiling(price / _slab) * _slab
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            ret = Math.Floor(price / _slab) * _slab
        End If
        Return ret
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Decimal)
        Dim ret As Tuple(Of Boolean, Decimal) = Nothing
        Dim lastExecutedOrder As Trade = _parentStrategy.GetLastSpecificTrades(candle, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Inprogress)
        If lastExecutedOrder Is Nothing Then
            Dim buyLevel As Decimal = GetSlabBasedLevel(currentTick.Open, Trade.TradeExecutionDirection.Buy)
            Dim sellLevel As Decimal = GetSlabBasedLevel(currentTick.Open, Trade.TradeExecutionDirection.Sell)
            If candle.High >= buyLevel AndAlso candle.Low <= sellLevel Then
                If direction = Trade.TradeExecutionDirection.Buy Then
                    ret = New Tuple(Of Boolean, Decimal)(True, buyLevel)
                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                    ret = New Tuple(Of Boolean, Decimal)(True, sellLevel)
                End If
            ElseIf candle.High >= buyLevel Then
                If direction = Trade.TradeExecutionDirection.Sell Then
                    ret = New Tuple(Of Boolean, Decimal)(True, sellLevel)
                End If
            ElseIf candle.Low <= sellLevel Then
                If direction = Trade.TradeExecutionDirection.Buy Then
                    ret = New Tuple(Of Boolean, Decimal)(True, buyLevel)
                End If
            End If
        Else
            If lastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                If lastExecutedOrder.EntryDirection <> direction Then
                    If lastExecutedOrder.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        ret = New Tuple(Of Boolean, Decimal)(True, lastExecutedOrder.PotentialStopLoss + lastExecutedOrder.StoplossBuffer)
                    ElseIf lastExecutedOrder.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        ret = New Tuple(Of Boolean, Decimal)(True, lastExecutedOrder.PotentialStopLoss - lastExecutedOrder.StoplossBuffer)
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class
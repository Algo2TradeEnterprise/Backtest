﻿Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class BelowFractalLowSingleStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public InitialCapital As Integer
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing

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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload, True)
        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentDayCandlePayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If currentDayCandlePayload IsNot Nothing AndAlso currentDayCandlePayload.PreviousCandlePayload IsNot Nothing Then
            Dim signals As List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer)) = GetSignalForEntry(currentDayCandlePayload)
            If signals IsNot Nothing AndAlso signals.Count > 0 Then
                For Each runningSignal In signals
                    Dim entryPrice As Decimal = runningSignal.Item2
                    Dim quantity As Integer = runningSignal.Item3
                    Dim iteration As Integer = runningSignal.Item4
                    Dim remarks As String = runningSignal.Item5
                    Dim fractalLow As Decimal = runningSignal.Item6
                    Dim multipler As Integer = runningSignal.Item7

                    Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = entryPrice,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = quantity,
                                                                .Stoploss = .EntryPrice - 1000000000,
                                                                .Target = .EntryPrice + 1000000000,
                                                                .Buffer = 0,
                                                                .SignalCandle = currentDayCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = iteration,
                                                                .Supporting2 = remarks,
                                                                .Supporting3 = fractalLow,
                                                                .Supporting4 = multipler
                                                            }

                    If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                    parameters.Add(parameter)
                Next
            End If
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

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
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

    Private Function GetSignalForEntry(ByVal candle As Payload) As List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer))
        Dim ret As List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer)) = Nothing
        Dim lastTrade As Trade = GetLastOrder(candle)
        Dim atr As Decimal = _atrPayload(candle.PreviousCandlePayload.PayloadDate)
        Dim iteration As Integer = 1
        Dim quantity As Integer = GetQuantity(iteration, candle.Open)
        Dim fractalLow As Decimal = _fractalLowPayload(candle.PreviousCandlePayload.PayloadDate)
        Dim entryPrice As Decimal = candle.Close
        If lastTrade Is Nothing Then
            fractalLow = _fractalLowPayload(candle.PayloadDate)
            If candle.Close < fractalLow Then
                If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer))
                ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer)(True, entryPrice, quantity, iteration, "First Trade", fractalLow, 1))
            End If
        Else
            Dim lastUsedFractal As Decimal = lastTrade.Supporting3
            If IsDifferentFractal(fractalLow, candle.PreviousCandlePayload.PayloadDate, lastUsedFractal, lastTrade.EntryTime) Then
                fractalLow = _fractalLowPayload(candle.PayloadDate)
                If candle.Close < fractalLow Then
                    If entryPrice < lastTrade.EntryPrice Then
                        If lastTrade.EntryPrice - entryPrice >= atr Then
                            If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer))
                            ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer)(True, entryPrice, quantity, iteration, "(NF)(Reset) Below last entry", fractalLow, 1))

                            'iteration = Val(lastTrade.Supporting1) + 1
                            'quantity = GetQuantity(iteration, entryPrice)
                            'If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer))
                            'ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer)(True, entryPrice, quantity, iteration, "(NF) Below last entry", fractalLow, 1))
                        End If
                    Else
                        If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer))
                        ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer)(True, entryPrice, quantity, iteration, "(NF)(Reset) Above last entry", fractalLow, 1))
                    End If
                End If
            Else
                Dim multiplier As Integer = lastTrade.Supporting4 * 2
                entryPrice = Utilities.Numbers.ConvertFloorCeling(lastTrade.EntryPrice - atr * multiplier, _parentStrategy.TickSize, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                If candle.Low <= entryPrice Then
                    iteration = Val(lastTrade.Supporting1) + 1
                    quantity = GetQuantity(iteration, entryPrice)
                    If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer))
                    ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer)(True, entryPrice, quantity, iteration, "Below last entry", fractalLow, multiplier))
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsDifferentFractal(ByVal currentFractal As Decimal, ByVal currentFratalTime As Date,
                                        ByVal lastUsedFractal As Decimal, ByVal lastusedFratalTime As Date) As Boolean
        Dim ret As Boolean = False
        If currentFractal <> lastUsedFractal Then
            ret = True
        Else
            For Each runningFractal In _fractalLowPayload.OrderByDescending(Function(x)
                                                                                Return x.Key
                                                                            End Function)
                If runningFractal.Key <= currentFratalTime AndAlso runningFractal.Key > lastusedFratalTime Then
                    If runningFractal.Value <> currentFractal Then
                        ret = True
                        Exit For
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Private Function GetQuantity(ByVal iterationNumber As Integer, ByVal price As Decimal) As Integer
        Dim capital As Decimal = _userInputs.InitialCapital * Math.Pow(2, iterationNumber - 1)
        Return _parentStrategy.CalculateQuantityFromInvestment(1, capital, price, Trade.TypeOfStock.Cash, False)
    End Function

    Private Function GetLastOrder(ByVal currentPayload As Payload) As Trade
        Dim ret As Trade = Nothing
        If currentPayload IsNot Nothing Then
            ret = Me._parentStrategy.GetLastExecutedTradeOfTheStock(currentPayload, Me._parentStrategy.TradeType)
        End If
        Return ret
    End Function
End Class
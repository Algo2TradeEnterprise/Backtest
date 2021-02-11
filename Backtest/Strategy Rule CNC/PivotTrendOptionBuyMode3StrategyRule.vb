﻿Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers
Imports Backtest.StrategyHelper

Public Class PivotTrendOptionBuyMode3StrategyRule
    Inherits StrategyRule

    Private _pivotTrendPayload As Dictionary(Of Date, Color) = Nothing
    Private _currentDayPivotTrends As Dictionary(Of Date, Color) = Nothing

    Public Sub New(ByVal canceller As CancellationTokenSource,
                    ByVal tradingDate As Date,
                    ByVal nextTradingDay As Date,
                    ByVal tradingSymbol As String,
                    ByVal lotSize As Integer,
                    ByVal parentStrategy As Strategy,
                    ByVal inputMinPayload As Dictionary(Of Date, Payload),
                    ByVal inputEODPayload As Dictionary(Of Date, Payload))
        MyBase.New(canceller, tradingDate, nextTradingDay, tradingSymbol, lotSize, parentStrategy, inputMinPayload, inputEODPayload)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.PivotHighLow.CalculatePivotHighLowTrend(4, 3, _SignalPayload, Nothing, Nothing, _pivotTrendPayload)
        If _pivotTrendPayload IsNot Nothing AndAlso _pivotTrendPayload.Count > 0 Then
            For Each runningPaylod In _pivotTrendPayload.OrderByDescending(Function(x)
                                                                               Return x.Key
                                                                           End Function)
                If runningPaylod.Key.Date = _TradingDate.Date Then
                    If _currentDayPivotTrends Is Nothing Then _currentDayPivotTrends = New Dictionary(Of Date, Color)
                    _currentDayPivotTrends.Add(runningPaylod.Key, runningPaylod.Value)
                Else
                    Exit For
                End If
            Next
            If _currentDayPivotTrends IsNot Nothing AndAlso _currentDayPivotTrends.Count > 0 Then
                For Each runningPaylod In _currentDayPivotTrends
                    _pivotTrendPayload.Remove(runningPaylod.Key)
                Next
            End If
        End If
    End Sub

    Public Overrides Async Function UpdateCollectionsIfRequiredAsync(currentTickTime As Date) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        If _ParentStrategy.SignalTimeFrame >= 375 Then
            If currentTickTime >= _TradeStartTime Then
                If Not _pivotTrendPayload.ContainsKey(_TradingDate) AndAlso _currentDayPivotTrends.ContainsKey(_TradingDate) Then
                    _pivotTrendPayload.Add(_TradingDate, _currentDayPivotTrends(_TradingDate))
                End If
            End If
        Else
            Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime, _ParentStrategy.SignalTimeFrame)
            If (currentTickTime >= currentMinute.AddMinutes(_ParentStrategy.SignalTimeFrame - 2) OrElse currentTickTime >= _TradeStartTime) AndAlso
                Not _pivotTrendPayload.ContainsKey(currentMinute) AndAlso _currentDayPivotTrends.ContainsKey(currentMinute) Then
                _pivotTrendPayload.Add(currentMinute, _currentDayPivotTrends(currentMinute))
            End If
        End If
    End Function

    Protected Overrides Function IsReverseSignalReceived(currentTrade As Trade, currentTick As Payload) As Tuple(Of Boolean, Payload)
        Dim ret As Tuple(Of Boolean, Payload) = Nothing
        Dim trend As Color = _pivotTrendPayload.LastOrDefault.Value
        If trend = Color.Red AndAlso currentTrade.SignalDirection = Trade.TradeDirection.Buy Then
            ret = New Tuple(Of Boolean, Payload)(True, _SignalPayload(_pivotTrendPayload.LastOrDefault.Key))
        ElseIf trend = Color.Green AndAlso currentTrade.SignalDirection = Trade.TradeDirection.Sell Then
            ret = New Tuple(Of Boolean, Payload)(True, _SignalPayload(_pivotTrendPayload.LastOrDefault.Key))
        End If
        Return ret
    End Function

    Protected Overrides Function IsFreshEntrySignalReceived(currentTickTime As Date) As Tuple(Of Boolean, Payload, Trade.TradeDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeDirection) = Nothing
        Dim potentialSignalCandle As Payload = _SignalPayload(_pivotTrendPayload.LastOrDefault.Key)
        Dim trend As Color = _pivotTrendPayload(potentialSignalCandle.PayloadDate)
        Dim previousTrend As Color = _pivotTrendPayload(potentialSignalCandle.PreviousCandlePayload.PayloadDate)
        If trend = Color.Green Then
            If previousTrend = Color.Red Then
                ret = New Tuple(Of Boolean, Payload, Trade.TradeDirection)(True, potentialSignalCandle, Trade.TradeDirection.Buy)
            Else
                Dim rolloverDay As Date = GetRolloverDay(trend, potentialSignalCandle.PayloadDate)
                If rolloverDay <> Date.MinValue Then
                    potentialSignalCandle = _SignalPayload(rolloverDay)
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeDirection)(True, potentialSignalCandle, Trade.TradeDirection.Buy)
                End If
            End If
        ElseIf trend = Color.Red Then
            If previousTrend = Color.Green Then
                ret = New Tuple(Of Boolean, Payload, Trade.TradeDirection)(True, potentialSignalCandle, Trade.TradeDirection.Sell)
            Else
                Dim rolloverDay As Date = GetRolloverDay(trend, potentialSignalCandle.PayloadDate)
                If rolloverDay <> Date.MinValue Then
                    potentialSignalCandle = _SignalPayload(rolloverDay)
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeDirection)(True, potentialSignalCandle, Trade.TradeDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetRolloverDay(ByVal currentTrend As Color, ByVal currentTime As Date) As Date
        Dim ret As Date = Date.MinValue
        For Each runningPayload In _SignalPayload.OrderByDescending(Function(x)
                                                                        Return x.Key
                                                                    End Function)
            If runningPayload.Value.PreviousCandlePayload IsNot Nothing AndAlso
                runningPayload.Value.PreviousCandlePayload.PayloadDate < currentTime Then
                Dim trend As Color = _pivotTrendPayload(runningPayload.Value.PreviousCandlePayload.PayloadDate)
                If trend <> currentTrend Then
                    ret = runningPayload.Key
                    Exit For
                End If
            End If
        Next
        Return ret
    End Function
End Class
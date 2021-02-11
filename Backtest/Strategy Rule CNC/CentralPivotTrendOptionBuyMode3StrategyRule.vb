Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers
Imports Backtest.StrategyHelper

Public Class CentralPivotTrendOptionBuyMode3StrategyRule
    Inherits StrategyRule

    Private _pivotPointsPayload As Dictionary(Of Date, PivotPoints) = Nothing
    Private _currentDayPivotPoints As Dictionary(Of Date, PivotPoints) = Nothing

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

        CalculatePivotPoints(_SignalPayload, _pivotPointsPayload)
        If _pivotPointsPayload IsNot Nothing AndAlso _pivotPointsPayload.Count > 0 Then
            For Each runningPaylod In _pivotPointsPayload.OrderByDescending(Function(x)
                                                                                Return x.Key
                                                                            End Function)
                If runningPaylod.Key.Date = _TradingDate.Date Then
                    If _currentDayPivotPoints Is Nothing Then _currentDayPivotPoints = New Dictionary(Of Date, PivotPoints)
                    _currentDayPivotPoints.Add(runningPaylod.Key, runningPaylod.Value)
                Else
                    Exit For
                End If
            Next
            If _currentDayPivotPoints IsNot Nothing AndAlso _currentDayPivotPoints.Count > 0 Then
                For Each runningPaylod In _currentDayPivotPoints
                    _pivotPointsPayload.Remove(runningPaylod.Key)
                Next
            End If
        End If
    End Sub

    Public Overrides Async Function UpdateCollectionsIfRequiredAsync(currentTickTime As Date) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        If _ParentStrategy.SignalTimeFrame >= 375 Then
            If currentTickTime >= _TradeStartTime Then
                If Not _pivotPointsPayload.ContainsKey(_TradingDate) AndAlso _currentDayPivotPoints.ContainsKey(_TradingDate) Then
                    _pivotPointsPayload.Add(_TradingDate, _currentDayPivotPoints(_TradingDate))
                End If
            End If
        Else
            Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime, _ParentStrategy.SignalTimeFrame)
            If (currentTickTime >= currentMinute.AddMinutes(_ParentStrategy.SignalTimeFrame - 2) OrElse currentTickTime >= _TradeStartTime) AndAlso
                Not _pivotPointsPayload.ContainsKey(currentMinute) AndAlso _currentDayPivotPoints.ContainsKey(currentMinute) Then
                _pivotPointsPayload.Add(currentMinute, _currentDayPivotPoints(currentMinute))
            End If
        End If
    End Function

    Protected Overrides Function IsReverseSignalReceived(currentTrade As Trade, currentTick As Payload) As Tuple(Of Boolean, Payload)
        Dim ret As Tuple(Of Boolean, Payload) = Nothing
        Dim pivot As PivotPoints = _pivotPointsPayload.LastOrDefault.Value
        Dim previousPivot As PivotPoints = _pivotPointsPayload(_SignalPayload(_pivotPointsPayload.LastOrDefault.Key).PreviousCandlePayload.PayloadDate)
        If currentTrade.SignalDirection = Trade.TradeDirection.Buy AndAlso pivot.Pivot < previousPivot.Pivot Then
            ret = New Tuple(Of Boolean, Payload)(True, _SignalPayload(_pivotPointsPayload.LastOrDefault.Key))
        ElseIf currentTrade.SignalDirection = Trade.TradeDirection.Sell AndAlso pivot.Pivot > previousPivot.Pivot Then
            ret = New Tuple(Of Boolean, Payload)(True, _SignalPayload(_pivotPointsPayload.LastOrDefault.Key))
        End If
        Return ret
    End Function

    Protected Overrides Function IsFreshEntrySignalReceived(currentTickTime As Date) As Tuple(Of Boolean, Payload, Trade.TradeDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeDirection) = Nothing
        Dim potentialSignalCandle As Payload = _SignalPayload(_pivotPointsPayload.LastOrDefault.Key)
        Dim pivot As PivotPoints = _pivotPointsPayload(potentialSignalCandle.PayloadDate)
        Dim previousPivot As PivotPoints = _pivotPointsPayload(potentialSignalCandle.PreviousCandlePayload.PayloadDate)
        If pivot.Pivot > previousPivot.Pivot Then
            If pivot.Pivot > previousPivot.Resistance1 Then
                ret = New Tuple(Of Boolean, Payload, Trade.TradeDirection)(True, potentialSignalCandle, Trade.TradeDirection.Buy)
            Else
                Dim rolloverDay As Date = GetRolloverDay(potentialSignalCandle.PayloadDate, Trade.TradeDirection.Buy)
                If rolloverDay <> Date.MinValue Then
                    potentialSignalCandle = _SignalPayload(rolloverDay)
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeDirection)(True, potentialSignalCandle, Trade.TradeDirection.Buy)
                End If
            End If
        ElseIf pivot.Pivot < previousPivot.Pivot Then
            If pivot.Pivot < previousPivot.Support1 Then
                ret = New Tuple(Of Boolean, Payload, Trade.TradeDirection)(True, potentialSignalCandle, Trade.TradeDirection.Sell)
            Else
                Dim rolloverDay As Date = GetRolloverDay(potentialSignalCandle.PayloadDate, Trade.TradeDirection.Sell)
                If rolloverDay <> Date.MinValue Then
                    potentialSignalCandle = _SignalPayload(rolloverDay)
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeDirection)(True, potentialSignalCandle, Trade.TradeDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetRolloverDay(ByVal currentTime As Date, ByVal direction As Trade.TradeDirection) As Date
        Dim ret As Date = Date.MinValue
        For Each runningPayload In _SignalPayload.OrderByDescending(Function(x)
                                                                        Return x.Key
                                                                    End Function)
            If runningPayload.Value.PreviousCandlePayload IsNot Nothing AndAlso
                runningPayload.Value.PreviousCandlePayload.PayloadDate < currentTime Then
                Dim pivot As PivotPoints = _pivotPointsPayload(runningPayload.Value.PayloadDate)
                Dim previousPivot As PivotPoints = _pivotPointsPayload(runningPayload.Value.PreviousCandlePayload.PayloadDate)
                If direction = Trade.TradeDirection.Buy Then
                    If pivot.Pivot > previousPivot.Resistance1 Then
                        ret = runningPayload.Key
                        Exit For
                    ElseIf pivot.Pivot < previousPivot.Pivot Then
                        Exit For
                    End If
                ElseIf direction = Trade.TradeDirection.Sell Then
                    If pivot.Pivot < previousPivot.Support1 Then
                        ret = runningPayload.Key
                        Exit For
                    ElseIf pivot.Pivot > previousPivot.Pivot Then
                        Exit For
                    End If
                End If
            End If
        Next
        Return ret
    End Function

#Region "Pivot Points Trend Calculation"
    Private Sub CalculatePivotPoints(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputPayload As Dictionary(Of Date, PivotPoints))
        If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
            For Each runningPayload In inputPayload
                Dim pivotPointsData As PivotPoints = New PivotPoints
                Dim curHigh As Decimal = runningPayload.Value.High
                Dim curLow As Decimal = runningPayload.Value.Low
                Dim curClose As Decimal = runningPayload.Value.Close

                pivotPointsData.Pivot = (curHigh + curLow + curClose) / 3
                pivotPointsData.Support1 = (2 * pivotPointsData.Pivot) - curHigh
                pivotPointsData.Resistance1 = (2 * pivotPointsData.Pivot) - curLow
                pivotPointsData.Support2 = pivotPointsData.Pivot - (curHigh - curLow)
                pivotPointsData.Resistance2 = pivotPointsData.Pivot + (curHigh - curLow)
                pivotPointsData.Support3 = pivotPointsData.Support2 - (curHigh - curLow)
                pivotPointsData.Resistance3 = pivotPointsData.Resistance2 + (curHigh - curLow)

                If outputPayload Is Nothing Then outputPayload = New Dictionary(Of Date, PivotPoints)
                outputPayload.Add(runningPayload.Key, pivotPointsData)
            Next
        End If
    End Sub
#End Region
End Class
Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers
Imports Backtest.StrategyHelper

Public Class HKMATrendOptionBuyMode3StrategyRule
    Inherits StrategyRule

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _smaPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _hkmaTrendPayload As Dictionary(Of Date, Color) = Nothing

    Private _currentDayHKPayloads As Dictionary(Of Date, Payload) = Nothing
    Private _currentDaySMAPayloads As Dictionary(Of Date, Decimal) = Nothing
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

        Indicator.HeikenAshi.ConvertToHeikenAshi(_SignalPayload, _hkPayload)
        Indicator.SMA.CalculateSMA(50, Payload.PayloadFields.Close, _hkPayload, _smaPayload)
        CalculateHKMATrend(_SignalPayload, _hkmaTrendPayload)

        If _hkPayload IsNot Nothing AndAlso _hkPayload.Count > 0 Then
            For Each runningPaylod In _hkPayload.OrderByDescending(Function(x)
                                                                       Return x.Key
                                                                   End Function)
                If runningPaylod.Key.Date = _TradingDate.Date Then
                    If _currentDayHKPayloads Is Nothing Then _currentDayHKPayloads = New Dictionary(Of Date, Payload)
                    _currentDayHKPayloads.Add(runningPaylod.Key, _hkPayload(runningPaylod.Key))
                    If _currentDaySMAPayloads Is Nothing Then _currentDaySMAPayloads = New Dictionary(Of Date, Decimal)
                    _currentDaySMAPayloads.Add(runningPaylod.Key, _smaPayload(runningPaylod.Key))
                    If _currentDayPivotTrends Is Nothing Then _currentDayPivotTrends = New Dictionary(Of Date, Color)
                    _currentDayPivotTrends.Add(runningPaylod.Key, _hkmaTrendPayload(runningPaylod.Key))
                Else
                    Exit For
                End If
            Next
            If _currentDayHKPayloads IsNot Nothing AndAlso _currentDayHKPayloads.Count > 0 Then
                For Each runningPaylod In _currentDayHKPayloads
                    _hkPayload.Remove(runningPaylod.Key)
                    _smaPayload.Remove(runningPaylod.Key)
                    _hkmaTrendPayload.Remove(runningPaylod.Key)
                Next
            End If
        End If


    End Sub

    Public Overrides Async Function UpdateCollectionsIfRequiredAsync(currentTickTime As Date) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        If _ParentStrategy.SignalTimeFrame >= 375 Then
            If currentTickTime >= _TradeStartTime Then
                If Not _hkPayload.ContainsKey(_TradingDate) AndAlso _currentDayHKPayloads.ContainsKey(_TradingDate) Then
                    _hkPayload.Add(_TradingDate, _currentDayHKPayloads(_TradingDate))
                    _smaPayload.Add(_TradingDate, _currentDaySMAPayloads(_TradingDate))
                    _hkmaTrendPayload.Add(_TradingDate, _currentDayPivotTrends(_TradingDate))
                End If
            End If
        Else
            Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime, _ParentStrategy.SignalTimeFrame)
            If (currentTickTime >= currentMinute.AddMinutes(_ParentStrategy.SignalTimeFrame - 2) OrElse currentTickTime >= _TradeStartTime) AndAlso
                Not _hkPayload.ContainsKey(currentMinute) AndAlso _currentDayHKPayloads.ContainsKey(currentMinute) Then
                _hkPayload.Add(currentMinute, _currentDayHKPayloads(currentMinute))
                _smaPayload.Add(currentMinute, _currentDaySMAPayloads(currentMinute))
                _hkmaTrendPayload.Add(currentMinute, _currentDayPivotTrends(currentMinute))
            End If
        End If
    End Function

    Protected Overrides Function IsReverseSignalReceived(currentTrade As Trade, currentTick As Payload) As Tuple(Of Boolean, Payload)
        Dim ret As Tuple(Of Boolean, Payload) = Nothing
        If _ParentStrategy.RuleSettings.TypeOfReverse = RuleEntities.ReverseType.Strong Then
            Dim trend As Color = _hkmaTrendPayload.LastOrDefault.Value
            If trend = Color.Red AndAlso currentTrade.SignalDirection = Trade.TradeDirection.Buy Then
                ret = New Tuple(Of Boolean, Payload)(True, _SignalPayload(_hkmaTrendPayload.LastOrDefault.Key))
            ElseIf trend = Color.Green AndAlso currentTrade.SignalDirection = Trade.TradeDirection.Sell Then
                ret = New Tuple(Of Boolean, Payload)(True, _SignalPayload(_hkmaTrendPayload.LastOrDefault.Key))
            End If
        Else
            Dim hkCandle As Payload = _hkPayload.LastOrDefault.Value
            Dim sma As Decimal = _smaPayload(hkCandle.PayloadDate)
            If currentTrade.SignalDirection = Trade.TradeDirection.Buy Then
                If hkCandle.CandleColor = Color.Red AndAlso hkCandle.Close < sma Then
                    ret = New Tuple(Of Boolean, Payload)(True, _SignalPayload(hkCandle.PayloadDate))
                End If
            ElseIf currentTrade.SignalDirection = Trade.TradeDirection.Sell Then
                If hkCandle.CandleColor = Color.Green AndAlso hkCandle.Close > sma Then
                    ret = New Tuple(Of Boolean, Payload)(True, _SignalPayload(hkCandle.PayloadDate))
                End If
            End If
        End If
        Return ret
    End Function

    Protected Overrides Function IsFreshEntrySignalReceived(currentTickTime As Date) As Tuple(Of Boolean, Payload, Trade.TradeDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeDirection) = Nothing
        Dim potentialSignalCandle As Payload = _SignalPayload(_hkmaTrendPayload.LastOrDefault.Key)
        Dim trend As Color = _hkmaTrendPayload(potentialSignalCandle.PayloadDate)
        Dim previousTrend As Color = _hkmaTrendPayload(potentialSignalCandle.PreviousCandlePayload.PayloadDate)
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
        If ret IsNot Nothing AndAlso ret.Item1 Then
            Dim hkCandle As Payload = _hkPayload.LastOrDefault.Value
            Dim sma As Decimal = _smaPayload(hkCandle.PayloadDate)
            If ret.Item3 = Trade.TradeDirection.Buy Then
                If hkCandle.Close < sma Then
                    ret = Nothing
                End If
            ElseIf ret.Item3 = Trade.TradeDirection.Sell Then
                If hkCandle.Close > sma Then
                    ret = Nothing
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
                Dim trend As Color = _hkmaTrendPayload(runningPayload.Value.PreviousCandlePayload.PayloadDate)
                If trend <> currentTrend Then
                    ret = runningPayload.Key
                    Exit For
                End If
            End If
        Next
        Return ret
    End Function

#Region "HK MA Trend Calculation"
    Private Sub CalculateHKMATrend(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputPayload As Dictionary(Of Date, Color))
        If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
            Dim hkPayload As Dictionary(Of Date, Payload) = Nothing
            Indicator.HeikenAshi.ConvertToHeikenAshi(inputPayload, hkPayload)
            Dim smaPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.SMA.CalculateSMA(50, Payload.PayloadFields.Close, hkPayload, smaPayload)

            Dim trend As Color = Color.White
            For Each runningPayload In hkPayload
                If runningPayload.Value.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                    If runningPayload.Value.Close > smaPayload(runningPayload.Key) Then
                        trend = Color.Green
                    End If
                ElseIf runningPayload.Value.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                    If runningPayload.Value.Close < smaPayload(runningPayload.Key) Then
                        trend = Color.Red
                    End If
                End If

                If outputPayload Is Nothing Then outputPayload = New Dictionary(Of Date, Color)
                outputPayload.Add(runningPayload.Key, trend)
            Next
        End If
    End Sub
#End Region
End Class
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class StochasticDivergenceStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public MaxLossPerTrade As Decimal
        Public TakeAnyHighLow As Boolean
    End Class
#End Region

    Private _smiPayload As Dictionary(Of Date, Decimal) = Nothing

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

        Indicator.SMI.CalculateSMI(10, 3, 3, 10, _signalPayload, _smiPayload, Nothing)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signal As Tuple(Of Boolean, Decimal, Payload, String) = GetEntrySignal(currentMinuteCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim lastTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy)
                    If lastTrade Is Nothing OrElse lastTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                        Dim signalCandle As Payload = signal.Item3
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                        Dim entryPrice As Decimal = signalCandle.High + buffer
                        Dim stoploss As Decimal = signal.Item2 - buffer
                        Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                        Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade * _userInputs.TargetMultiplier), Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)

                        parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"),
                                .Supporting2 = signal.Item4
                            }
                    End If
                End If
            End If
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signal As Tuple(Of Boolean, Decimal, Payload, String) = GetEntrySignal(currentMinuteCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim lastTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell)
                    If lastTrade Is Nothing OrElse lastTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                        Dim signalCandle As Payload = signal.Item3
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                        Dim entryPrice As Decimal = signalCandle.Low - buffer
                        Dim stoploss As Decimal = signal.Item2 + buffer
                        Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                        Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade * _userInputs.TargetMultiplier), Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Cash)

                        parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"),
                                    .Supporting2 = signal.Item4
                                }
                    End If
                End If
            End If
        End If
        Dim parameterList As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            If parameterList Is Nothing Then parameterList = New List(Of PlaceOrderParameters)
            parameterList.Add(parameter1)
        End If
        If parameter2 IsNot Nothing Then
            If parameterList Is Nothing Then parameterList = New List(Of PlaceOrderParameters)
            parameterList.Add(parameter2)
        End If
        If parameterList IsNot Nothing AndAlso parameterList.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameterList)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            'Dim signal As Tuple(Of Boolean, Decimal, Payload, String) = GetEntrySignal(currentMinuteCandlePayload, currentTick, currentTrade.EntryDirection)
            'If signal IsNot Nothing AndAlso signal.Item1 Then
            '    If currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
            '        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            '    End If
            'End If
            Dim lastMountain As Tuple(Of Date, Date) = GetMoutainCandles(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate, currentTrade.EntryDirection)
            If lastMountain IsNot Nothing AndAlso lastMountain.Item1 <> Date.MinValue Then
                If currentTrade.SignalCandle.PayloadDate <> lastMountain.Item2 Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
            If ret Is Nothing Then
                Dim smi As Decimal = _smiPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso smi > 40 Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso smi < -40 Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Decimal, Payload, String)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, String) = Nothing
        Dim lastMountain As Tuple(Of Date, Date) = GetMoutainCandles(candle.PreviousCandlePayload.PayloadDate, direction)
        If IsSignalValid(lastMountain.Item2, candle.PreviousCandlePayload.PayloadDate, direction) AndAlso
            lastMountain IsNot Nothing AndAlso lastMountain.Item1 <> Date.MinValue AndAlso lastMountain.Item1.Date = _tradingDate.Date Then
            Dim preLastMountain As Tuple(Of Date, Date) = GetMoutainCandles(lastMountain.Item1, direction)
            If preLastMountain IsNot Nothing AndAlso preLastMountain.Item1 <> Date.MinValue Then
                If direction = Trade.TradeExecutionDirection.Sell Then
                    Dim lastMountainHigh As Decimal = _signalPayload.Max(Function(x)
                                                                             If x.Key >= lastMountain.Item1 AndAlso x.Key <= lastMountain.Item2 Then
                                                                                 Return x.Value.High
                                                                             Else
                                                                                 Return Decimal.MinValue
                                                                             End If
                                                                         End Function)
                    Dim lastMountainSMIHigh As Decimal = _smiPayload.Max(Function(x)
                                                                             If x.Key >= lastMountain.Item1 AndAlso x.Key <= lastMountain.Item2 Then
                                                                                 Return x.Value
                                                                             Else
                                                                                 Return Decimal.MinValue
                                                                             End If
                                                                         End Function)

                    Dim preLastMountainHigh As Decimal = _signalPayload.Max(Function(x)
                                                                                If x.Key >= preLastMountain.Item1 AndAlso x.Key <= preLastMountain.Item2 Then
                                                                                    Return x.Value.High
                                                                                Else
                                                                                    Return Decimal.MinValue
                                                                                End If
                                                                            End Function)
                    Dim preLastMountainSMIHigh As Decimal = _smiPayload.Max(Function(x)
                                                                                If x.Key >= preLastMountain.Item1 AndAlso x.Key <= preLastMountain.Item2 Then
                                                                                    Return x.Value
                                                                                Else
                                                                                    Return Decimal.MinValue
                                                                                End If
                                                                            End Function)
                    If Not ((lastMountainHigh > preLastMountainHigh AndAlso lastMountainSMIHigh > preLastMountainSMIHigh) OrElse
                        (lastMountainHigh < preLastMountainHigh AndAlso lastMountainSMIHigh < preLastMountainSMIHigh)) Then
                        Dim signalCandle As Payload = _signalPayload(lastMountain.Item2)
                        ret = New Tuple(Of Boolean, Decimal, Payload, String)(True, lastMountainHigh, signalCandle, String.Format("{0},{1},{2},{3},{4},{5},{6},{7}",
                                                                                                                                  preLastMountain.Item1.ToString("dd-MMM-yyyy HH:mm:ss"),
                                                                                                                                  preLastMountain.Item2.ToString("dd-MMM-yyyy HH:mm:ss"),
                                                                                                                                  preLastMountainHigh,
                                                                                                                                  preLastMountainSMIHigh,
                                                                                                                                  lastMountain.Item1.ToString("dd-MMM-yyyy HH:mm:ss"),
                                                                                                                                  lastMountain.Item2.ToString("dd-MMM-yyyy HH:mm:ss"),
                                                                                                                                  lastMountainHigh,
                                                                                                                                  lastMountainSMIHigh))
                    End If
                ElseIf direction = Trade.TradeExecutionDirection.Buy Then
                    Dim lastMountainLow As Decimal = _signalPayload.Min(Function(x)
                                                                            If x.Key >= lastMountain.Item1 AndAlso x.Key <= lastMountain.Item2 Then
                                                                                Return x.Value.Low
                                                                            Else
                                                                                Return Decimal.MaxValue
                                                                            End If
                                                                        End Function)
                    Dim lastMountainSMILow As Decimal = _smiPayload.Min(Function(x)
                                                                            If x.Key >= lastMountain.Item1 AndAlso x.Key <= lastMountain.Item2 Then
                                                                                Return x.Value
                                                                            Else
                                                                                Return Decimal.MaxValue
                                                                            End If
                                                                        End Function)

                    Dim preLastMountainLow As Decimal = _signalPayload.Min(Function(x)
                                                                               If x.Key >= preLastMountain.Item1 AndAlso x.Key <= preLastMountain.Item2 Then
                                                                                   Return x.Value.Low
                                                                               Else
                                                                                   Return Decimal.MaxValue
                                                                               End If
                                                                           End Function)
                    Dim preLastMountainSMILow As Decimal = _smiPayload.Min(Function(x)
                                                                               If x.Key >= preLastMountain.Item1 AndAlso x.Key <= preLastMountain.Item2 Then
                                                                                   Return x.Value
                                                                               Else
                                                                                   Return Decimal.MaxValue
                                                                               End If
                                                                           End Function)
                    If Not ((lastMountainLow > preLastMountainLow AndAlso lastMountainSMILow > preLastMountainSMILow) OrElse
                        (lastMountainLow < preLastMountainLow AndAlso lastMountainSMILow < preLastMountainSMILow)) Then
                        Dim signalCandle As Payload = _signalPayload(lastMountain.Item2)
                        ret = New Tuple(Of Boolean, Decimal, Payload, String)(True, lastMountainLow, signalCandle, String.Format("{0},{1},{2},{3},{4},{5},{6},{7}",
                                                                                                                                  preLastMountain.Item1.ToString("dd-MMM-yyyy HH:mm:ss"),
                                                                                                                                  preLastMountain.Item2.ToString("dd-MMM-yyyy HH:mm:ss"),
                                                                                                                                  preLastMountainLow,
                                                                                                                                  preLastMountainSMILow,
                                                                                                                                  lastMountain.Item1.ToString("dd-MMM-yyyy HH:mm:ss"),
                                                                                                                                  lastMountain.Item2.ToString("dd-MMM-yyyy HH:mm:ss"),
                                                                                                                                  lastMountainLow,
                                                                                                                                  lastMountainSMILow))
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetMoutainCandles(ByVal beforeThisTime As Date, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Date, Date)
        Dim ret As Tuple(Of Date, Date) = Nothing
        If _smiPayload IsNot Nothing AndAlso _smiPayload.Count > 0 Then
            Dim checkingPayload As IEnumerable(Of KeyValuePair(Of Date, Decimal)) = _smiPayload.Where(Function(x)
                                                                                                          Return x.Key <= beforeThisTime
                                                                                                      End Function)
            If checkingPayload IsNot Nothing AndAlso checkingPayload.Count > 0 Then
                Dim firstCandleTime As Date = Date.MinValue
                Dim middleCandleTime As Date = Date.MinValue
                Dim lastCandleTime As Date = Date.MinValue
                For Each runningPayload In checkingPayload.OrderByDescending(Function(x)
                                                                                 Return x.Key
                                                                             End Function)
                    If direction = Trade.TradeExecutionDirection.Sell Then
                        If Not _userInputs.TakeAnyHighLow Then
                            If middleCandleTime <> Date.MinValue AndAlso _smiPayload(middleCandleTime) < 40 Then
                                firstCandleTime = middleCandleTime
                                middleCandleTime = Date.MinValue
                                lastCandleTime = Date.MinValue
                            End If
                        End If
                        If firstCandleTime = Date.MinValue Then
                            firstCandleTime = runningPayload.Key
                        Else
                            If middleCandleTime = Date.MinValue Then
                                If runningPayload.Value > _smiPayload(firstCandleTime) Then
                                    middleCandleTime = runningPayload.Key
                                Else
                                    firstCandleTime = runningPayload.Key
                                End If
                            Else
                                If runningPayload.Value = _smiPayload(middleCandleTime) Then
                                    middleCandleTime = runningPayload.Key
                                ElseIf runningPayload.Value > _smiPayload(middleCandleTime) Then
                                    firstCandleTime = middleCandleTime
                                    middleCandleTime = runningPayload.Key
                                ElseIf runningPayload.Value < _smiPayload(middleCandleTime) Then
                                    lastCandleTime = runningPayload.Key
                                    ret = New Tuple(Of Date, Date)(lastCandleTime, firstCandleTime)
                                    Exit For
                                End If
                            End If
                        End If
                    ElseIf direction = Trade.TradeExecutionDirection.Buy Then
                        If Not _userInputs.TakeAnyHighLow Then
                            If middleCandleTime <> Date.MinValue AndAlso _smiPayload(middleCandleTime) > -40 Then
                                firstCandleTime = middleCandleTime
                                middleCandleTime = Date.MinValue
                                lastCandleTime = Date.MinValue
                            End If
                        End If
                        If firstCandleTime = Date.MinValue Then
                            firstCandleTime = runningPayload.Key
                        Else
                            If middleCandleTime = Date.MinValue Then
                                If runningPayload.Value < _smiPayload(firstCandleTime) Then
                                    middleCandleTime = runningPayload.Key
                                Else
                                    firstCandleTime = runningPayload.Key
                                End If
                            Else
                                If runningPayload.Value = _smiPayload(middleCandleTime) Then
                                    middleCandleTime = runningPayload.Key
                                ElseIf runningPayload.Value < _smiPayload(middleCandleTime) Then
                                    firstCandleTime = middleCandleTime
                                    middleCandleTime = runningPayload.Key
                                ElseIf runningPayload.Value > _smiPayload(middleCandleTime) Then
                                    lastCandleTime = runningPayload.Key
                                    ret = New Tuple(Of Date, Date)(lastCandleTime, firstCandleTime)
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
        Return ret
    End Function

    Private Function IsSignalValid(ByVal startTime As Date, ByVal currentTime As Date, ByVal direction As Trade.TradeExecutionDirection) As Boolean
        Dim ret As Boolean = False
        If startTime <> Date.MinValue Then
            For Each runningSMI In _smiPayload
                If runningSMI.Key >= startTime AndAlso runningSMI.Key <= currentTime Then
                    ret = True
                    If direction = Trade.TradeExecutionDirection.Buy AndAlso runningSMI.Value > 40 Then
                        ret = False
                        Exit For
                    ElseIf direction = Trade.TradeExecutionDirection.Sell AndAlso runningSMI.Value < -40 Then
                        ret = False
                        Exit For
                    End If
                End If
            Next
        End If
        Return ret
    End Function
End Class
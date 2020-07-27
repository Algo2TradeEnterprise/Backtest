Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class FibonacciBacktestStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public Multiplier As Decimal
    End Class
#End Region

    Private ReadOnly _entryLevel As Decimal = 38.2 / 100
    Private ReadOnly _target1Level As Decimal = 61.8 / 100
    Private ReadOnly _target2Level As Decimal = 100 / 100

    Private _avgMvmnt As Decimal = Decimal.MinValue

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

        Dim eodPayload As Dictionary(Of Date, Payload) = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_Futures, _tradingSymbol, _tradingDate.AddDays(-15), _tradingDate)
        If eodPayload IsNot Nothing AndAlso eodPayload.Count > 6 AndAlso eodPayload.ContainsKey(_tradingDate.Date) Then
            Dim hlPayload As Dictionary(Of Date, Payload) = Nothing
            For Each runningPayload In eodPayload
                If hlPayload Is Nothing Then hlPayload = New Dictionary(Of Date, Payload)
                Dim hl As Payload = New Payload(Payload.CandleDataSource.Calculated)
                hl.Close = runningPayload.Value.High - runningPayload.Value.Low
                hlPayload.Add(runningPayload.Key, hl)
            Next

            Dim smaPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.SMA.CalculateSMA(5, Payload.PayloadFields.Close, hlPayload, smaPayload)
            If smaPayload IsNot Nothing AndAlso smaPayload.Count > 0 Then
                Dim currentDayPayload As Payload = eodPayload(_tradingDate.Date)
                _avgMvmnt = smaPayload(currentDayPayload.PreviousCandlePayload.PayloadDate) * _userInputs.Multiplier
            End If
        Else
            Throw New ApplicationException("Previous 5 days data not available")
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso _avgMvmnt <> Decimal.MinValue AndAlso
            Not IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload) Then

            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item2
            End If

            If signalCandle IsNot Nothing Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim stoploss As Decimal = _signalPayload(signal.Item4).Low
                    Dim target1 As Decimal = stoploss + ConvertFloorCeling(_target1Level * _avgMvmnt, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim quantity As Integer = Me.LotSize

                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target1,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signal.Item4.ToString("HH:mm:ss"),
                                .Supporting2 = 1
                            }

                    Dim target2 As Decimal = stoploss + ConvertFloorCeling(_target2Level * _avgMvmnt, _parentStrategy.TickSize, RoundOfType.Floor)
                    parameter2 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target2,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signal.Item4.ToString("HH:mm:ss"),
                                .Supporting2 = 2
                            }

                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim stoploss As Decimal = _signalPayload(signal.Item4).High
                    Dim target1 As Decimal = stoploss - ConvertFloorCeling(_target1Level * _avgMvmnt, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim quantity As Integer = Me.LotSize

                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target1,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signal.Item4.ToString("HH:mm:ss"),
                                .Supporting2 = 1
                            }

                    Dim target2 As Decimal = stoploss - ConvertFloorCeling(_target2Level * _avgMvmnt, _parentStrategy.TickSize, RoundOfType.Floor)
                    parameter2 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target2,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signal.Item4.ToString("HH:mm:ss"),
                                .Supporting2 = 2
                            }
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
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryDirection <> signal.Item3 Then
                    ret = New Tuple(Of Boolean, String)(True, "Opposite Direction Signal")
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date) = Nothing
        If candle IsNot Nothing Then
            Dim lastTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(candle, Trade.TypeOfTrade.MIS)
            If lastTrade IsNot Nothing Then
                If lastTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    Dim highestPayload As Payload = GetHighestHighPayload(lastTrade.EntryTime, currentTick.PayloadDate)
                    If highestPayload IsNot Nothing Then
                        Dim sellLevel As Decimal = highestPayload.High - ConvertFloorCeling(_avgMvmnt * _entryLevel, _parentStrategy.TickSize, RoundOfType.Floor)
                        If currentTick.Open <= sellLevel Then
                            ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date)(True, candle, Trade.TradeExecutionDirection.Sell, highestPayload.PayloadDate)
                        End If
                    End If
                ElseIf lastTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    Dim lowestPayload As Payload = GetLowestLowPayload(lastTrade.EntryTime, currentTick.PayloadDate)
                    If lowestPayload IsNot Nothing Then
                        Dim buyLevel As Decimal = lowestPayload.Low + ConvertFloorCeling(_avgMvmnt * _entryLevel, _parentStrategy.TickSize, RoundOfType.Floor)
                        If currentTick.Open >= buyLevel Then
                            ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date)(True, candle, Trade.TradeExecutionDirection.Buy, lowestPayload.PayloadDate)
                        End If
                    End If
                End If
            Else
                Dim highestPayload As Payload = GetHighestHighPayload(_tradingDate, currentTick.PayloadDate)
                Dim lowestPayload As Payload = GetLowestLowPayload(_tradingDate, currentTick.PayloadDate)
                Dim buyLevel As Decimal = lowestPayload.Low + ConvertFloorCeling(_avgMvmnt * _entryLevel, _parentStrategy.TickSize, RoundOfType.Floor)
                Dim sellLevel As Decimal = highestPayload.High - ConvertFloorCeling(_avgMvmnt * _entryLevel, _parentStrategy.TickSize, RoundOfType.Floor)
                If currentTick.Open >= buyLevel Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date)(True, candle, Trade.TradeExecutionDirection.Buy, lowestPayload.PayloadDate)
                ElseIf currentTick.Open <= sellLevel Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date)(True, candle, Trade.TradeExecutionDirection.Sell, highestPayload.PayloadDate)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestHighPayload(ByVal startTime As Date, ByVal endTime As Date) As Payload
        Dim ret As Payload = Nothing
        For Each runningPayload In _signalPayload
            If runningPayload.Key > startTime AndAlso runningPayload.Key < endTime Then
                If ret Is Nothing Then ret = runningPayload.Value
                If runningPayload.Value.High >= ret.High Then
                    ret = runningPayload.Value
                End If
            End If
        Next
        Return ret
    End Function

    Private Function GetLowestLowPayload(ByVal startTime As Date, ByVal endTime As Date) As Payload
        Dim ret As Payload = Nothing
        For Each runningPayload In _signalPayload
            If runningPayload.Key > startTime AndAlso runningPayload.Key < endTime Then
                If ret Is Nothing Then ret = runningPayload.Value
                If runningPayload.Value.Low <= ret.Low Then
                    ret = runningPayload.Value
                End If
            End If
        Next
        Return ret
    End Function

    Private Function IsAnyTradeOfTheStockTargetReached(ByVal currentMinutePayload As Payload) As Boolean
        Dim ret As Boolean = False
        Dim completeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
        If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then
            Dim targetTrades As List(Of Trade) = completeTrades.FindAll(Function(x)
                                                                            Return x.ExitCondition = Trade.TradeExitCondition.Target AndAlso
                                                                            x.AdditionalTrade = False AndAlso x.Supporting2 = "2"
                                                                        End Function)
            If targetTrades IsNot Nothing AndAlso targetTrades.Count > 0 Then
                ret = True
            End If
        End If
        Return ret
    End Function
End Class
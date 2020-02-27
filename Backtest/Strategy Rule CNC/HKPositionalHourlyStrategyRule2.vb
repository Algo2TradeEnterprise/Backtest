Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HKPositionalHourlyStrategyRule2
    Inherits StrategyRule

    Private _hkPayload As Dictionary(Of Date, Payload)
    Private _entryForContractRolloverDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Private _exitForContractRollover As Boolean
    Private _previousSymbolPayloads As Dictionary(Of Date, Payload)

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

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)

        Dim previousTradingDay As Date = _parentStrategy.Cmn.GetPreviousTradingDay(_parentStrategy.DatabaseTable, _tradingDate)
        If previousTradingDay <> Date.MinValue Then
            Dim rawInstrument As String = _tradingSymbol.Remove(_tradingSymbol.Count - 8)
            Dim previousTradingDaySymbol As String = _parentStrategy.Cmn.GetCurrentTradingSymbol(_parentStrategy.DatabaseTable, previousTradingDay, rawInstrument)
            If previousTradingDaySymbol IsNot Nothing AndAlso previousTradingDaySymbol.ToUpper <> _tradingSymbol.ToUpper Then
                _exitForContractRollover = True
                _previousSymbolPayloads = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(_parentStrategy.DatabaseTable, previousTradingDaySymbol, _tradingDate, _tradingDate)
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _exitForContractRollover Then ExitForContractRollover(currentTick)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) AndAlso Not _parentStrategy.IsTradeActive(currentTick, _parentStrategy.TradeType) Then
            Dim signalCandle As Payload = Nothing

            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String) = GetSignalForEntry(currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item3
                If signalCandle IsNot Nothing Then
                    'Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                    Dim buffer As Decimal = 1
                    If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                        parameter = New PlaceOrderParameters With {
                                        .EntryPrice = currentTick.Open,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = Me.LotSize,
                                        .Stoploss = ConvertFloorCeling(signalCandle.Open - buffer, _parentStrategy.TickSize, RoundOfType.Floor),
                                        .Target = .EntryPrice + 1000000,
                                        .Buffer = buffer,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting1 = signal.Item5
                                    }
                    ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                        parameter = New PlaceOrderParameters With {
                                        .EntryPrice = currentTick.Open,
                                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                        .Quantity = Me.LotSize,
                                        .Stoploss = ConvertFloorCeling(signalCandle.Open + buffer, _parentStrategy.TickSize, RoundOfType.Floor),
                                        .Target = .EntryPrice - 1000000,
                                        .Buffer = buffer,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting1 = signal.Item5
                                    }
                    End If
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            _entryForContractRolloverDirection = Trade.TradeExecutionDirection.None
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim currentHKPayload As Payload = _hkPayload(currentMinuteCandlePayload.PayloadDate)
            If currentTrade.PotentialStopLoss <> currentHKPayload.Open - currentTrade.StoplossBuffer Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, currentHKPayload.Open - currentTrade.StoplossBuffer, String.Format("Modified at {0}",
                                                                                                        currentTick.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss")))
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function GetSignalForEntry(ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String) = Nothing
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim currentHKPayload As Payload = _hkPayload(currentMinuteCandlePayload.PayloadDate)
        If currentTick.PayloadDate >= currentHKPayload.PayloadDate.AddMinutes(59) OrElse
            currentTick.PayloadDate >= New Date(currentTick.PayloadDate.Year, currentTick.PayloadDate.Month, currentTick.PayloadDate.Day, 15, 29, 0) Then
            If currentHKPayload.Open - 1 <= currentHKPayload.Low Then
                ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, currentTick.Open, currentHKPayload, Trade.TradeExecutionDirection.Buy, "Normal Entry")
            ElseIf currentHKPayload.Open + 1 >= currentHKPayload.High Then
                ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, currentTick.Open, currentHKPayload, Trade.TradeExecutionDirection.Sell, "Normal Entry")
            End If
        ElseIf _entryForContractRolloverDirection <> Trade.TradeExecutionDirection.None Then
            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, currentTick.Open, currentHKPayload, _entryForContractRolloverDirection, "Contract Rollover Entry")
        End If
        Return ret
    End Function

    Private Sub ExitForContractRollover(ByVal currentTick As Payload)
        If _exitForContractRollover Then
            Dim currentCandlePayload As Payload = Nothing
            If _previousSymbolPayloads IsNot Nothing AndAlso _previousSymbolPayloads.Count > 0 Then
                currentCandlePayload = _previousSymbolPayloads(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _previousSymbolPayloads))
            End If
            If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
                currentCandlePayload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            End If
            If currentCandlePayload IsNot Nothing Then
                Dim potentialRuleExitTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
                If potentialRuleExitTrades IsNot Nothing AndAlso potentialRuleExitTrades.Count > 0 Then
                    For Each runningExitTrade In potentialRuleExitTrades
                        Dim exitTick As Payload = New Payload(Payload.CandleDataSource.Calculated) With {
                                                        .Open = currentCandlePayload.Open,
                                                        .Low = currentCandlePayload.Open,
                                                        .High = currentCandlePayload.Open,
                                                        .Close = currentCandlePayload.Open,
                                                        .PayloadDate = currentCandlePayload.PayloadDate,
                                                        .TradingSymbol = currentCandlePayload.TradingSymbol
                                                    }
                        _parentStrategy.ExitTradeByForce(runningExitTrade, exitTick, "Contract Rollover Exit")
                    Next
                    _exitForContractRollover = False
                    _entryForContractRolloverDirection = potentialRuleExitTrades.LastOrDefault.EntryDirection
                End If
            End If
        End If
    End Sub

End Class

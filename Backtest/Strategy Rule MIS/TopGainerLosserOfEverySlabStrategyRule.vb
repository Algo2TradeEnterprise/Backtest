Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class TopGainerLosserOfEverySlabStrategyRule
    Inherits StrategyRule

    Private _firstTimeEntry As Boolean = True
    Private _lastDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None

    Public Direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Public GainLossPerData As Dictionary(Of Date, Decimal) = Nothing
    Public DummyCandle As Payload = Nothing

    Private ReadOnly _changePer As Decimal = 0.1
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal controller As Integer,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, controller, canceller)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Dim currentDayPayload As Dictionary(Of Date, Payload) = Nothing
        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = _tradingDate.Date Then
                If currentDayPayload Is Nothing Then currentDayPayload = New Dictionary(Of Date, Payload)
                currentDayPayload.Add(runningPayload.Key, runningPayload.Value)
            End If
        Next
        If currentDayPayload IsNot Nothing AndAlso currentDayPayload.Count > 0 Then
            Dim firstCandle As Payload = currentDayPayload.FirstOrDefault.Value
            Dim previousClose As Decimal = firstCandle.PreviousCandlePayload.Close
            For Each runningPayload In currentDayPayload
                Dim gainLossPercentage As Decimal = ((runningPayload.Value.Close - previousClose) / previousClose) * 100
                If GainLossPerData Is Nothing Then GainLossPerData = New Dictionary(Of Date, Decimal)
                GainLossPerData.Add(runningPayload.Key, gainLossPercentage)
            Next
        End If
    End Sub

    Public Overrides Sub CompletePairProcessing()
        MyBase.CompletePairProcessing()
        If Me.Controller Then
            If Me.DependentInstrument IsNot Nothing AndAlso Me.DependentInstrument.Count > 0 Then
                For Each runningInstrument In Me.DependentInstrument
                    runningInstrument.EligibleToTakeTrade = False
                Next
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Me.DummyCandle = currentTick
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If Me.ForceTakeTrade AndAlso Not Controller Then
            Dim quantity As Integer = _parentStrategy.CalculateQuantityFromInvestment(Me.LotSize, 5000, currentTick.Open, Trade.TypeOfStock.Cash, False)
            If Me.Direction = Trade.TradeExecutionDirection.Buy Then
                Dim parameter As PlaceOrderParameters = Nothing
                parameter = New PlaceOrderParameters With {
                                .EntryPrice = currentTick.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - 10000000,
                                .Target = .EntryPrice + 10000000,
                                .Buffer = 0,
                                .SignalCandle = currentMinuteCandlePayload,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = "Force Entry"
                            }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                Me.ForceTakeTrade = False
            ElseIf Me.Direction = Trade.TradeExecutionDirection.Sell Then
                Dim parameter As PlaceOrderParameters = Nothing
                parameter = New PlaceOrderParameters With {
                                .EntryPrice = currentTick.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + 10000000,
                                .Target = .EntryPrice - 10000000,
                                .Buffer = 0,
                                .SignalCandle = currentMinuteCandlePayload,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = "Force Entry"
                            }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                Me.ForceTakeTrade = False
            End If
        ElseIf Controller Then
            Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
            Dim parameter As PlaceOrderParameters = Nothing
            If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
                currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
                If _firstTimeEntry Then
                    If Me.GainLossPerData IsNot Nothing AndAlso
                        Me.GainLossPerData.ContainsKey(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) Then
                        If Me.GainLossPerData(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) > _changePer Then
                            _lastDirection = Trade.TradeExecutionDirection.Buy
                            Dim topGainers As List(Of TopGainerLosserOfEverySlabStrategyRule) = GetTopGainerStocks(currentMinuteCandlePayload, 2)
                            If topGainers IsNot Nothing AndAlso topGainers.Count > 0 Then
                                For Each runningGainer In topGainers
                                    runningGainer.Direction = Trade.TradeExecutionDirection.Buy
                                    runningGainer.ForceTakeTrade = True
                                    runningGainer.EligibleToTakeTrade = True
                                Next
                            End If
                            Dim topLosser As List(Of TopGainerLosserOfEverySlabStrategyRule) = GetTopLosserStocks(currentMinuteCandlePayload, 1)
                            If topLosser IsNot Nothing AndAlso topLosser.Count > 0 Then
                                For Each runningLosser In topLosser
                                    runningLosser.Direction = Trade.TradeExecutionDirection.Sell
                                    runningLosser.ForceTakeTrade = True
                                    runningLosser.EligibleToTakeTrade = True
                                Next
                            End If
                            _firstTimeEntry = False
                        ElseIf Me.GainLossPerData(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) < _changePer * -1 Then
                            _lastDirection = Trade.TradeExecutionDirection.Sell
                            Dim topGainers As List(Of TopGainerLosserOfEverySlabStrategyRule) = GetTopGainerStocks(currentMinuteCandlePayload, 1)
                            If topGainers IsNot Nothing AndAlso topGainers.Count > 0 Then
                                For Each runningGainer In topGainers
                                    runningGainer.Direction = Trade.TradeExecutionDirection.Buy
                                    runningGainer.ForceTakeTrade = True
                                    runningGainer.EligibleToTakeTrade = True
                                Next
                            End If
                            Dim topLosser As List(Of TopGainerLosserOfEverySlabStrategyRule) = GetTopLosserStocks(currentMinuteCandlePayload, 2)
                            If topLosser IsNot Nothing AndAlso topLosser.Count > 0 Then
                                For Each runningLosser In topLosser
                                    runningLosser.Direction = Trade.TradeExecutionDirection.Sell
                                    runningLosser.ForceTakeTrade = True
                                    runningLosser.EligibleToTakeTrade = True
                                Next
                            End If
                            _firstTimeEntry = False
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Me.DummyCandle = currentTick
        If Me.ForceCancelTrade AndAlso Not Controller Then
            If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                ret = New Tuple(Of Boolean, String)(True, "Force Exit")
                Me.ForceCancelTrade = False
            End If
        ElseIf Controller Then
            If _lastDirection <> Trade.TradeExecutionDirection.None Then
                Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                If _lastDirection = Trade.TradeExecutionDirection.Buy Then
                    If Me.GainLossPerData IsNot Nothing AndAlso
                        Me.GainLossPerData.ContainsKey(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) Then
                        If Me.GainLossPerData(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) < _changePer * -1 Then
                            _lastDirection = Trade.TradeExecutionDirection.Sell
                            Dim activeInstrument As List(Of TopGainerLosserOfEverySlabStrategyRule) = GetActiveInstruments()
                            If activeInstrument IsNot Nothing AndAlso activeInstrument.Count > 0 Then
                                For Each runningInstrument In activeInstrument
                                    runningInstrument.ForceCancelTrade = True
                                Next
                            End If
                            Dim topGainers As List(Of TopGainerLosserOfEverySlabStrategyRule) = GetTopGainerStocks(currentMinuteCandlePayload, 1)
                            If topGainers IsNot Nothing AndAlso topGainers.Count > 0 Then
                                For Each runningGainer In topGainers
                                    runningGainer.Direction = Trade.TradeExecutionDirection.Buy
                                    runningGainer.ForceTakeTrade = True
                                    runningGainer.EligibleToTakeTrade = True
                                Next
                            End If
                            Dim topLosser As List(Of TopGainerLosserOfEverySlabStrategyRule) = GetTopLosserStocks(currentMinuteCandlePayload, 2)
                            If topLosser IsNot Nothing AndAlso topLosser.Count > 0 Then
                                For Each runningLosser In topLosser
                                    runningLosser.Direction = Trade.TradeExecutionDirection.Sell
                                    runningLosser.ForceTakeTrade = True
                                    runningLosser.EligibleToTakeTrade = True
                                Next
                            End If
                        End If
                    End If
                ElseIf _lastDirection = Trade.TradeExecutionDirection.Sell Then
                    If Me.GainLossPerData IsNot Nothing AndAlso
                        Me.GainLossPerData.ContainsKey(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) Then
                        If Me.GainLossPerData(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) > _changePer Then
                            _lastDirection = Trade.TradeExecutionDirection.Buy
                            Dim activeInstrument As List(Of TopGainerLosserOfEverySlabStrategyRule) = GetActiveInstruments()
                            If activeInstrument IsNot Nothing AndAlso activeInstrument.Count > 0 Then
                                For Each runningInstrument In activeInstrument
                                    runningInstrument.ForceCancelTrade = True
                                Next
                            End If
                            Dim topGainers As List(Of TopGainerLosserOfEverySlabStrategyRule) = GetTopGainerStocks(currentMinuteCandlePayload, 2)
                            If topGainers IsNot Nothing AndAlso topGainers.Count > 0 Then
                                For Each runningGainer In topGainers
                                    runningGainer.Direction = Trade.TradeExecutionDirection.Buy
                                    runningGainer.ForceTakeTrade = True
                                    runningGainer.EligibleToTakeTrade = True
                                Next
                            End If
                            Dim topLosser As List(Of TopGainerLosserOfEverySlabStrategyRule) = GetTopLosserStocks(currentMinuteCandlePayload, 1)
                            If topLosser IsNot Nothing AndAlso topLosser.Count > 0 Then
                                For Each runningLosser In topLosser
                                    runningLosser.Direction = Trade.TradeExecutionDirection.Sell
                                    runningLosser.ForceTakeTrade = True
                                    runningLosser.EligibleToTakeTrade = True
                                Next
                            End If
                        End If
                    End If
                End If
            End If
        End If
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

    Private Function GetTopGainerStocks(ByVal currentMinuteCandle As Payload, ByVal numberOfStocks As Integer) As List(Of TopGainerLosserOfEverySlabStrategyRule)
        Dim ret As List(Of TopGainerLosserOfEverySlabStrategyRule) = Nothing
        If Me.DependentInstrument IsNot Nothing AndAlso Me.DependentInstrument.Count > 0 Then
            Dim gainLossOfTheMinute As Dictionary(Of String, Decimal) = Nothing
            For Each runningInstrument As TopGainerLosserOfEverySlabStrategyRule In Me.DependentInstrument
                If runningInstrument.GainLossPerData IsNot Nothing AndAlso
                    runningInstrument.GainLossPerData.ContainsKey(currentMinuteCandle.PreviousCandlePayload.PayloadDate) Then
                    If gainLossOfTheMinute Is Nothing Then gainLossOfTheMinute = New Dictionary(Of String, Decimal)
                    gainLossOfTheMinute.Add(runningInstrument.TradingSymbol, runningInstrument.GainLossPerData(currentMinuteCandle.PreviousCandlePayload.PayloadDate))
                End If
            Next
            If gainLossOfTheMinute IsNot Nothing AndAlso gainLossOfTheMinute.Count > 0 Then
                For Each runningStock In gainLossOfTheMinute.OrderByDescending(Function(x)
                                                                                   Return x.Value
                                                                               End Function)
                    If ret Is Nothing Then ret = New List(Of TopGainerLosserOfEverySlabStrategyRule)
                    ret.Add(GetStrategyRuleByName(runningStock.Key))

                    If ret.Count = numberOfStocks Then Exit For
                Next
            End If
        End If
        Return ret
    End Function

    Private Function GetTopLosserStocks(ByVal currentMinuteCandle As Payload, ByVal numberOfStocks As Integer) As List(Of TopGainerLosserOfEverySlabStrategyRule)
        Dim ret As List(Of TopGainerLosserOfEverySlabStrategyRule) = Nothing
        If Me.DependentInstrument IsNot Nothing AndAlso Me.DependentInstrument.Count > 0 Then
            Dim gainLossOfTheMinute As Dictionary(Of String, Decimal) = Nothing
            For Each runningInstrument As TopGainerLosserOfEverySlabStrategyRule In Me.DependentInstrument
                If runningInstrument.GainLossPerData IsNot Nothing AndAlso
                    runningInstrument.GainLossPerData.ContainsKey(currentMinuteCandle.PreviousCandlePayload.PayloadDate) Then
                    If gainLossOfTheMinute Is Nothing Then gainLossOfTheMinute = New Dictionary(Of String, Decimal)
                    gainLossOfTheMinute.Add(runningInstrument.TradingSymbol, runningInstrument.GainLossPerData(currentMinuteCandle.PreviousCandlePayload.PayloadDate))
                End If
            Next
            If gainLossOfTheMinute IsNot Nothing AndAlso gainLossOfTheMinute.Count > 0 Then
                For Each runningStock In gainLossOfTheMinute.OrderBy(Function(x)
                                                                         Return x.Value
                                                                     End Function)
                    If ret Is Nothing Then ret = New List(Of TopGainerLosserOfEverySlabStrategyRule)
                    ret.Add(GetStrategyRuleByName(runningStock.Key))

                    If ret.Count = numberOfStocks Then Exit For
                Next
            End If
        End If
        Return ret
    End Function

    Private Function GetStrategyRuleByName(ByVal tradingSymbol As String) As TopGainerLosserOfEverySlabStrategyRule
        Dim ret As TopGainerLosserOfEverySlabStrategyRule = Nothing
        If Me.DependentInstrument IsNot Nothing AndAlso Me.DependentInstrument.Count > 0 Then
            For Each runningInstrument In Me.DependentInstrument
                If runningInstrument.TradingSymbol.ToUpper = tradingSymbol.ToUpper Then
                    ret = runningInstrument
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function

    Private Function GetActiveInstruments() As List(Of TopGainerLosserOfEverySlabStrategyRule)
        Dim ret As List(Of TopGainerLosserOfEverySlabStrategyRule) = Nothing
        If Me.DependentInstrument IsNot Nothing AndAlso Me.DependentInstrument.Count > 0 Then
            For Each runningInstrument As TopGainerLosserOfEverySlabStrategyRule In Me.DependentInstrument
                If runningInstrument.DummyCandle IsNot Nothing AndAlso
                    _parentStrategy.IsTradeActive(runningInstrument.DummyCandle, Trade.TypeOfTrade.MIS) Then
                    If ret Is Nothing Then ret = New List(Of TopGainerLosserOfEverySlabStrategyRule)
                    ret.Add(runningInstrument)
                End If
            Next
        End If
        Return ret
    End Function
End Class
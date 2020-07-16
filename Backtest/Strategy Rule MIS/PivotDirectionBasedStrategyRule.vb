Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PivotDirectionBasedStrategyRule
    Inherits StrategyRule

    Public DummyCandle As Payload = Nothing
    Public DirectionToTrade As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None

    Private ReadOnly _controller As Boolean
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal canceller As CancellationTokenSource,
                   ByVal controller As Boolean)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, canceller)

        _controller = controller
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
    End Sub

    Public Overrides Sub CompletePairProcessing()
        MyBase.CompletePairProcessing()

        If _controller Then
            Dim firstCandleOfTheDay As Payload = Nothing
            For Each runningPayload In _signalPayload
                If runningPayload.Key.Date = _tradingDate.Date Then
                    If runningPayload.Value.PreviousCandlePayload IsNot Nothing AndAlso
                        runningPayload.Value.PayloadDate.Date <> runningPayload.Value.PreviousCandlePayload.PayloadDate.Date Then
                        firstCandleOfTheDay = runningPayload.Value
                        Exit For
                    End If
                End If
            Next
            If firstCandleOfTheDay IsNot Nothing Then
                Dim pivotPayload As Dictionary(Of Date, PivotPoints) = Nothing
                Indicator.Pivots.CalculatePivots(_signalPayload, pivotPayload)

                If pivotPayload IsNot Nothing AndAlso pivotPayload.ContainsKey(firstCandleOfTheDay.PayloadDate) AndAlso
                    pivotPayload.ContainsKey(firstCandleOfTheDay.PreviousCandlePayload.PayloadDate) Then
                    If pivotPayload(firstCandleOfTheDay.PayloadDate).Pivot > pivotPayload(firstCandleOfTheDay.PreviousCandlePayload.PayloadDate).Pivot Then
                        Me.DirectionToTrade = Trade.TradeExecutionDirection.Buy
                        CType(Me.AnotherPairInstrument, PivotDirectionBasedStrategyRule).DirectionToTrade = Trade.TradeExecutionDirection.Sell
                    Else
                        Me.DirectionToTrade = Trade.TradeExecutionDirection.Sell
                        CType(Me.AnotherPairInstrument, PivotDirectionBasedStrategyRule).DirectionToTrade = Trade.TradeExecutionDirection.Sell
                    End If
                End If
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Me.DummyCandle = currentTick
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If DirectionToTrade = Trade.TradeExecutionDirection.Buy AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                parameter = New PlaceOrderParameters With {
                            .EntryPrice = currentTick.Open,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = Me.LotSize,
                            .Stoploss = .EntryPrice - 1000000000000,
                            .Target = .EntryPrice + 1000000000000,
                            .Buffer = 0,
                            .SignalCandle = currentMinuteCandlePayload,
                            .OrderType = Trade.TypeOfOrder.Market
                        }
            End If
            If DirectionToTrade = Trade.TradeExecutionDirection.Sell AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                parameter = New PlaceOrderParameters With {
                            .EntryPrice = currentTick.Open,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = Me.LotSize,
                            .Stoploss = .EntryPrice + 1000000000000,
                            .Target = .EntryPrice - 1000000000000,
                            .Buffer = 0,
                            .SignalCandle = currentMinuteCandlePayload,
                            .OrderType = Trade.TypeOfOrder.Market
                        }
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
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

    Public Shared Function GetTodaysStock(ByVal tradingDate As Date, ByVal dt As DataTable, ByVal cmn As Common) As Dictionary(Of String, StockDetails)
        Dim ret As Dictionary(Of String, StockDetails) = Nothing
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim stockList As List(Of StockDetails) = Nothing
            Dim mainStock As StockDetails = Nothing
            For i = 1 To dt.Rows.Count - 1
                Dim rowDate As Date = dt.Rows(i)(0)
                If rowDate.Date = tradingDate.Date Then
                    Dim tradingSymbol As String = dt.Rows(i).Item(1)
                    Dim controller As Boolean = False
                    If tradingSymbol.EndsWith("FUT") Then
                        controller = True
                    End If
                    Dim instrumentName As String = Nothing
                    If tradingSymbol.Contains("FUT") Then
                        instrumentName = tradingSymbol.Remove(tradingSymbol.Count - 8)
                    Else
                        instrumentName = tradingSymbol
                    End If
                    Dim detailsOfStock As StockDetails = New StockDetails With
                                {.StockName = instrumentName,
                                 .TradingSymbol = tradingSymbol,
                                 .LotSize = dt.Rows(i).Item(2),
                                 .EligibleToTakeTrade = True,
                                 .Controller = controller,
                                 .OptionStock = Not controller,
                                 .StockType = Trade.TypeOfStock.Futures}

                    If controller Then mainStock = detailsOfStock

                    If stockList Is Nothing Then stockList = New List(Of StockDetails)
                    stockList.Add(detailsOfStock)
                End If
            Next
            If mainStock IsNot Nothing Then
                Dim intradayPayload As Dictionary(Of Date, Payload) = cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures, mainStock.TradingSymbol, tradingDate.AddDays(-7), tradingDate)
                If intradayPayload IsNot Nothing AndAlso intradayPayload.Count > 0 Then
                    Dim firstCandleOfTheDay As Payload = Nothing
                    For Each runningPayload In intradayPayload
                        If runningPayload.Key.Date = tradingDate.Date Then
                            If runningPayload.Value.PreviousCandlePayload IsNot Nothing AndAlso
                                runningPayload.Value.PayloadDate.Date <> runningPayload.Value.PreviousCandlePayload.PayloadDate.Date Then
                                firstCandleOfTheDay = runningPayload.Value
                                Exit For
                            End If
                        End If
                    Next
                    If firstCandleOfTheDay IsNot Nothing Then
                        Dim pivotPayload As Dictionary(Of Date, PivotPoints) = Nothing
                        Indicator.Pivots.CalculatePivots(intradayPayload, pivotPayload)

                        If pivotPayload IsNot Nothing AndAlso pivotPayload.ContainsKey(firstCandleOfTheDay.PayloadDate) AndAlso
                            pivotPayload.ContainsKey(firstCandleOfTheDay.PreviousCandlePayload.PayloadDate) Then
                            Dim optionSuffix As String = Nothing
                            If pivotPayload(firstCandleOfTheDay.PayloadDate).Pivot > pivotPayload(firstCandleOfTheDay.PreviousCandlePayload.PayloadDate).Pivot Then
                                optionSuffix = "CE"
                            Else
                                optionSuffix = "PE"
                            End If
                            If optionSuffix IsNot Nothing Then
                                For Each runningStock In stockList
                                    If runningStock.OptionStock AndAlso runningStock.TradingSymbol.EndsWith(optionSuffix) Then
                                        Dim strikeOption As String = runningStock.TradingSymbol.Substring(14)
                                        If IsNumeric(strikeOption.Substring(0, strikeOption.Count - 2)) Then
                                            Dim strikePrice As Decimal = Val(strikeOption.Substring(0, strikeOption.Count - 2))
                                            If optionSuffix = "CE" Then
                                                If strikePrice <= firstCandleOfTheDay.Close AndAlso strikePrice + 100 > firstCandleOfTheDay.Close Then
                                                    ret = New Dictionary(Of String, StockDetails)
                                                    ret.Add(mainStock.StockName, mainStock)
                                                    ret.Add(runningStock.StockName, runningStock)
                                                    Exit For
                                                End If
                                            ElseIf optionSuffix = "PE" Then
                                                If strikePrice >= firstCandleOfTheDay.Close AndAlso strikePrice - 100 < firstCandleOfTheDay.Close Then
                                                    ret = New Dictionary(Of String, StockDetails)
                                                    ret.Add(mainStock.StockName, mainStock)
                                                    ret.Add(runningStock.StockName, runningStock)
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class
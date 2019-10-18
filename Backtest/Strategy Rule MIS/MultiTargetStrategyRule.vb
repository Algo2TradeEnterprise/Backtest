Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class MultiTargetStrategyRule
    Inherits StrategyRule

    Private ReadOnly _maxLossPercentage As Decimal = 35
    Private ReadOnly _minimumTargetPL As Decimal = 100

    Private _EODPayload As Dictionary(Of Date, Payload) = Nothing

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
        Dim dbTable As Common.DataBaseTable = Common.DataBaseTable.None
        Select Case Me._parentStrategy.DatabaseTable
            Case Common.DataBaseTable.Intraday_Cash
                dbTable = Common.DataBaseTable.EOD_Cash
            Case Common.DataBaseTable.Intraday_Futures
                dbTable = Common.DataBaseTable.EOD_Futures
        End Select
        _EODPayload = Me._parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(dbTable, _tradingSymbol, _tradingDate.AddDays(-10), _tradingDate)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day,
                                              _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > Math.Abs(_parentStrategy.OverAllLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandleSatisfied As Tuple(Of Boolean, TradeDetails) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                    signalCandleSatisfied.Item2.BuyEntry <> Decimal.MinValue AndAlso currentTick.Open >= signalCandleSatisfied.Item2.BuyEntry Then
                    Dim buffer As Decimal = Me._parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2.BuyEntry, RoundOfType.Floor)
                    Dim lastTargetEntryRequired As Boolean = False
                    If signalCandleSatisfied.Item2.BuyEntry < signalCandleSatisfied.Item2.BuyTarget1 AndAlso
                        IsTargetSatisfied(Trade.TradeExecutionDirection.Buy, signalCandleSatisfied.Item2.BuyEntry, signalCandleSatisfied.Item2.BuyTarget1) Then
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signalCandleSatisfied.Item2.BuyEntry,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = _lotSize,
                                                                .Stoploss = signalCandleSatisfied.Item2.BuyStoploss,
                                                                .Target = signalCandleSatisfied.Item2.BuyTarget1,
                                                                .Buffer = buffer,
                                                                .SignalCandle = currentMinuteCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = "Target 1",
                                                                .Supporting2 = signalCandleSatisfied.Item2.BuyLevel = signalCandleSatisfied.Item2.BuyEntry
                                                            }
                        If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                        parameters.Add(parameter)
                    Else
                        lastTargetEntryRequired = True
                    End If
                    If signalCandleSatisfied.Item2.BuyEntry < signalCandleSatisfied.Item2.BuyTarget2 AndAlso
                        IsTargetSatisfied(Trade.TradeExecutionDirection.Buy, signalCandleSatisfied.Item2.BuyEntry, signalCandleSatisfied.Item2.BuyTarget2) Then
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signalCandleSatisfied.Item2.BuyEntry,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = _lotSize,
                                                                .Stoploss = signalCandleSatisfied.Item2.BuyStoploss,
                                                                .Target = signalCandleSatisfied.Item2.BuyTarget2,
                                                                .Buffer = buffer,
                                                                .SignalCandle = currentMinuteCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = "Target 2",
                                                                .Supporting2 = signalCandleSatisfied.Item2.BuyLevel = signalCandleSatisfied.Item2.BuyEntry
                                                            }
                        If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                        parameters.Add(parameter)
                    End If
                    If signalCandleSatisfied.Item2.BuyEntry < signalCandleSatisfied.Item2.BuyTarget3 AndAlso
                        IsTargetSatisfied(Trade.TradeExecutionDirection.Buy, signalCandleSatisfied.Item2.BuyEntry, signalCandleSatisfied.Item2.BuyTarget3) Then
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signalCandleSatisfied.Item2.BuyEntry,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = _lotSize,
                                                                .Stoploss = signalCandleSatisfied.Item2.BuyStoploss,
                                                                .Target = signalCandleSatisfied.Item2.BuyTarget3,
                                                                .Buffer = buffer,
                                                                .SignalCandle = currentMinuteCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = "Target 3",
                                                                .Supporting2 = signalCandleSatisfied.Item2.BuyLevel = signalCandleSatisfied.Item2.BuyEntry
                                                            }
                        If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                        parameters.Add(parameter)
                    End If
                    If signalCandleSatisfied.Item2.BuyEntry < signalCandleSatisfied.Item2.BuyTarget4 AndAlso
                        IsTargetSatisfied(Trade.TradeExecutionDirection.Buy, signalCandleSatisfied.Item2.BuyEntry, signalCandleSatisfied.Item2.BuyTarget4) Then
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signalCandleSatisfied.Item2.BuyEntry,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = _lotSize,
                                                                .Stoploss = signalCandleSatisfied.Item2.BuyStoploss,
                                                                .Target = signalCandleSatisfied.Item2.BuyTarget4,
                                                                .Buffer = buffer,
                                                                .SignalCandle = currentMinuteCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = "Target 4",
                                                                .Supporting2 = signalCandleSatisfied.Item2.BuyLevel = signalCandleSatisfied.Item2.BuyEntry
                                                            }
                        If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                        parameters.Add(parameter)
                    End If
                    If lastTargetEntryRequired AndAlso signalCandleSatisfied.Item2.BuyEntry < signalCandleSatisfied.Item2.BuyTarget5 AndAlso
                        IsTargetSatisfied(Trade.TradeExecutionDirection.Buy, signalCandleSatisfied.Item2.BuyEntry, signalCandleSatisfied.Item2.BuyTarget5) Then
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signalCandleSatisfied.Item2.BuyEntry,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = _lotSize,
                                                                .Stoploss = signalCandleSatisfied.Item2.BuyStoploss,
                                                                .Target = signalCandleSatisfied.Item2.BuyTarget5,
                                                                .Buffer = buffer,
                                                                .SignalCandle = currentMinuteCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = "Target 5",
                                                                .Supporting2 = signalCandleSatisfied.Item2.BuyLevel = signalCandleSatisfied.Item2.BuyEntry
                                                            }
                        If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                        parameters.Add(parameter)
                    End If
                ElseIf Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                    signalCandleSatisfied.Item2.SellEntry <> Decimal.MinValue AndAlso currentTick.Open <= signalCandleSatisfied.Item2.SellEntry Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2.SellEntry, RoundOfType.Floor)
                    Dim lastTargetEntryRequired As Boolean = False
                    If signalCandleSatisfied.Item2.SellEntry > signalCandleSatisfied.Item2.SellTarget1 AndAlso
                        IsTargetSatisfied(Trade.TradeExecutionDirection.Sell, signalCandleSatisfied.Item2.SellEntry, signalCandleSatisfied.Item2.SellTarget1) Then
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signalCandleSatisfied.Item2.SellEntry,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                                                .Quantity = _lotSize,
                                                                .Stoploss = signalCandleSatisfied.Item2.SellStoploss,
                                                                .Target = signalCandleSatisfied.Item2.SellTarget1,
                                                                .Buffer = buffer,
                                                                .SignalCandle = currentMinuteCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = "Target 1",
                                                                .Supporting2 = signalCandleSatisfied.Item2.SellLevel = signalCandleSatisfied.Item2.SellEntry
                                                            }
                        If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                        parameters.Add(parameter)
                    Else
                        lastTargetEntryRequired = True
                    End If
                    If signalCandleSatisfied.Item2.SellEntry > signalCandleSatisfied.Item2.SellTarget2 AndAlso
                        IsTargetSatisfied(Trade.TradeExecutionDirection.Sell, signalCandleSatisfied.Item2.SellEntry, signalCandleSatisfied.Item2.SellTarget2) Then
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signalCandleSatisfied.Item2.SellEntry,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                                                .Quantity = _lotSize,
                                                                .Stoploss = signalCandleSatisfied.Item2.SellStoploss,
                                                                .Target = signalCandleSatisfied.Item2.SellTarget2,
                                                                .Buffer = buffer,
                                                                .SignalCandle = currentMinuteCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = "Target 2",
                                                                .Supporting2 = signalCandleSatisfied.Item2.SellLevel = signalCandleSatisfied.Item2.SellEntry
                                                            }
                        If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                        parameters.Add(parameter)
                    End If
                    If signalCandleSatisfied.Item2.SellEntry > signalCandleSatisfied.Item2.SellTarget3 AndAlso
                        IsTargetSatisfied(Trade.TradeExecutionDirection.Sell, signalCandleSatisfied.Item2.SellEntry, signalCandleSatisfied.Item2.SellTarget3) Then
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signalCandleSatisfied.Item2.SellEntry,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                                                .Quantity = _lotSize,
                                                                .Stoploss = signalCandleSatisfied.Item2.SellStoploss,
                                                                .Target = signalCandleSatisfied.Item2.SellTarget3,
                                                                .Buffer = buffer,
                                                                .SignalCandle = currentMinuteCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = "Target 3",
                                                                .Supporting2 = signalCandleSatisfied.Item2.SellLevel = signalCandleSatisfied.Item2.SellEntry
                                                            }
                        If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                        parameters.Add(parameter)
                    End If
                    If signalCandleSatisfied.Item2.SellEntry > signalCandleSatisfied.Item2.SellTarget4 AndAlso
                        IsTargetSatisfied(Trade.TradeExecutionDirection.Sell, signalCandleSatisfied.Item2.SellEntry, signalCandleSatisfied.Item2.SellTarget4) Then
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signalCandleSatisfied.Item2.SellEntry,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                                                .Quantity = _lotSize,
                                                                .Stoploss = signalCandleSatisfied.Item2.SellStoploss,
                                                                .Target = signalCandleSatisfied.Item2.SellTarget4,
                                                                .Buffer = buffer,
                                                                .SignalCandle = currentMinuteCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = "Target 4",
                                                                .Supporting2 = signalCandleSatisfied.Item2.SellLevel = signalCandleSatisfied.Item2.SellEntry
                                                            }
                        If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                        parameters.Add(parameter)
                    End If
                    If lastTargetEntryRequired AndAlso signalCandleSatisfied.Item2.SellEntry > signalCandleSatisfied.Item2.SellTarget5 AndAlso
                        IsTargetSatisfied(Trade.TradeExecutionDirection.Sell, signalCandleSatisfied.Item2.SellEntry, signalCandleSatisfied.Item2.SellTarget5) Then
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signalCandleSatisfied.Item2.SellEntry,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                                                .Quantity = _lotSize,
                                                                .Stoploss = signalCandleSatisfied.Item2.SellStoploss,
                                                                .Target = signalCandleSatisfied.Item2.SellTarget5,
                                                                .Buffer = buffer,
                                                                .SignalCandle = currentMinuteCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = "Target 5",
                                                                .Supporting2 = signalCandleSatisfied.Item2.SellLevel = signalCandleSatisfied.Item2.SellEntry
                                                            }
                        If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                        parameters.Add(parameter)
                    End If
                End If
            End If
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)

            'If _parentStrategy.StockMaxProfitPercentagePerDay <> Decimal.MaxValue AndAlso Me.MaxProfitOfThisStock = Decimal.MaxValue Then
            '    Me.MaxProfitOfThisStock = _parentStrategy.CalculatePL(currentTick.TradingSymbol, parameter.EntryPrice, ConvertFloorCeling(parameter.EntryPrice + parameter.EntryPrice * _parentStrategy.StockMaxProfitPercentagePerDay / 100, _parentStrategy.TickSize, RoundOfType.Celing), parameter.Quantity, _lotSize, _parentStrategy.StockType)
            'End If
            'If _parentStrategy.StockMaxLossPercentagePerDay <> Decimal.MinValue AndAlso Me.MaxLossOfThisStock = Decimal.MinValue Then
            '    Me.MaxLossOfThisStock = _parentStrategy.CalculatePL(currentTick.TradingSymbol, parameter.EntryPrice, ConvertFloorCeling(parameter.EntryPrice - parameter.EntryPrice * _parentStrategy.StockMaxLossPercentagePerDay / 100, _parentStrategy.TickSize, RoundOfType.Celing), parameter.Quantity, _lotSize, _parentStrategy.StockType)
            'End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
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

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, TradeDetails)
        Dim ret As Tuple(Of Boolean, TradeDetails) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim previousDayPayload As Payload = Nothing
            Dim currentDayPayload As Payload = Nothing
            If _EODPayload IsNot Nothing AndAlso _EODPayload.ContainsKey(_tradingDate) Then
                currentDayPayload = _EODPayload(_tradingDate)
                previousDayPayload = currentDayPayload.PreviousCandlePayload
            End If
            If previousDayPayload IsNot Nothing AndAlso currentDayPayload IsNot Nothing Then
                Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day,
                                                      _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
                If _signalPayload IsNot Nothing AndAlso _signalPayload.ContainsKey(tradeStartTime) Then
                    Dim currentPayload As Payload = _signalPayload(tradeStartTime)
                    Dim dayHigh As Decimal = _signalPayload.Max(Function(x)
                                                                    If x.Key.Date = _tradingDate.Date AndAlso x.Key < tradeStartTime Then
                                                                        Return x.Value.High
                                                                    Else
                                                                        Return Decimal.MinValue
                                                                    End If
                                                                End Function)
                    Dim dayLow As Decimal = _signalPayload.Min(Function(x)
                                                                   If x.Key.Date = _tradingDate.Date AndAlso x.Key < tradeStartTime Then
                                                                       Return x.Value.Low
                                                                   Else
                                                                       Return Decimal.MaxValue
                                                                   End If
                                                               End Function)
                    Dim dayOpen As Decimal = currentDayPayload.Open
                    Dim ltp As Decimal = currentPayload.Open

                    Dim pp4 As Decimal = (ltp + dayOpen + dayHigh + dayLow) / 4
                    Dim pp3 As Decimal = (previousDayPayload.High + previousDayPayload.Low + previousDayPayload.Close) / 3
                    Dim pHL As Decimal = previousDayPayload.High - previousDayPayload.Low

                    Dim tradeEntryDetails As TradeDetails = New TradeDetails
                    tradeEntryDetails.BuyLevel = ConvertFloorCeling((previousDayPayload.High + previousDayPayload.Low + previousDayPayload.Close) / 3, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.BuyStoploss = ConvertFloorCeling(pp3 - 0.51 * pHL, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.BuyTarget1 = tradeEntryDetails.BuyLevel + ConvertFloorCeling(0.236 * pHL, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.BuyTarget2 = tradeEntryDetails.BuyLevel + ConvertFloorCeling(0.382 * pHL, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.BuyTarget3 = tradeEntryDetails.BuyLevel + ConvertFloorCeling(0.5 * pHL, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.BuyTarget4 = tradeEntryDetails.BuyLevel + ConvertFloorCeling(0.618 * pHL, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.BuyTarget5 = tradeEntryDetails.BuyLevel + ConvertFloorCeling(0.764 * pHL, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.SellLevel = ConvertFloorCeling(pp4 - 0.236 * pHL, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.SellStoploss = ConvertFloorCeling(pp4, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.SellTarget1 = tradeEntryDetails.SellLevel - ConvertFloorCeling(0.236 * pHL, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.SellTarget2 = tradeEntryDetails.SellLevel - ConvertFloorCeling(0.382 * pHL, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.SellTarget3 = tradeEntryDetails.SellLevel - ConvertFloorCeling(0.5 * pHL, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.SellTarget4 = tradeEntryDetails.SellLevel - ConvertFloorCeling(0.618 * pHL, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    tradeEntryDetails.SellTarget5 = tradeEntryDetails.SellLevel - ConvertFloorCeling(0.764 * pHL, Me._parentStrategy.TickSize, RoundOfType.Celing)

                    If dayHigh < tradeEntryDetails.BuyLevel Then
                        tradeEntryDetails.BuyEntry = tradeEntryDetails.BuyLevel
                    ElseIf dayHigh < tradeEntryDetails.BuyTarget4 Then
                        tradeEntryDetails.BuyEntry = dayHigh + Me._parentStrategy.CalculateBuffer(dayHigh, RoundOfType.Floor)
                    Else
                        tradeEntryDetails.BuyEntry = Decimal.MinValue
                    End If

                    If dayLow > tradeEntryDetails.SellLevel Then
                        tradeEntryDetails.SellEntry = tradeEntryDetails.SellLevel
                    ElseIf dayLow > tradeEntryDetails.SellTarget4 Then
                        tradeEntryDetails.SellEntry = dayLow - Me._parentStrategy.CalculateBuffer(dayLow, RoundOfType.Floor)
                    Else
                        tradeEntryDetails.SellEntry = Decimal.MinValue
                    End If

                    If tradeEntryDetails.BuyEntry <> Decimal.MinValue Then
                        Dim stoplossPL As Decimal = Me._parentStrategy.CalculatePL(_tradingSymbol, tradeEntryDetails.BuyEntry, tradeEntryDetails.BuyStoploss, _lotSize, _lotSize, Me._parentStrategy.StockType)
                        Dim requiredCapital As Decimal = tradeEntryDetails.BuyEntry * _lotSize / Me._parentStrategy.MarginMultiplier
                        If Math.Abs(stoplossPL) > requiredCapital * _maxLossPercentage / 100 Then
                            tradeEntryDetails.BuyEntry = Decimal.MinValue
                        End If
                    End If

                    If tradeEntryDetails.SellEntry <> Decimal.MinValue Then
                        Dim stoplossPL As Decimal = Me._parentStrategy.CalculatePL(_tradingSymbol, tradeEntryDetails.SellStoploss, tradeEntryDetails.SellEntry, _lotSize, _lotSize, Me._parentStrategy.StockType)
                        Dim requiredCapital As Decimal = tradeEntryDetails.SellEntry * _lotSize / Me._parentStrategy.MarginMultiplier
                        If Math.Abs(stoplossPL) > requiredCapital * _maxLossPercentage / 100 Then
                            tradeEntryDetails.SellEntry = Decimal.MinValue
                        End If
                    End If

                    If tradeEntryDetails.BuyEntry = Decimal.MinValue AndAlso tradeEntryDetails.SellEntry = Decimal.MinValue Then
                        Me.EligibleToTakeTrade = False
                    End If

                    ret = New Tuple(Of Boolean, TradeDetails)(True, tradeEntryDetails)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsTargetSatisfied(ByVal direction As Trade.TradeExecutionDirection, ByVal entryPrice As Decimal, ByVal exitPrice As Decimal) As Boolean
        Dim ret As Boolean = False
        Dim pl As Decimal = Decimal.MinValue
        If direction = Trade.TradeExecutionDirection.Buy Then
            pl = Me._parentStrategy.CalculatePL(_tradingSymbol, entryPrice, exitPrice, _lotSize, _lotSize, Me._parentStrategy.StockType)
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            pl = Me._parentStrategy.CalculatePL(_tradingSymbol, exitPrice, entryPrice, _lotSize, _lotSize, Me._parentStrategy.StockType)
        End If
        If pl >= _minimumTargetPL Then
            ret = True
        End If
        Return ret
    End Function

    Private Class TradeDetails
        Public BuyLevel As Decimal
        Public BuyEntry As Decimal
        Public BuyStoploss As Decimal
        Public BuyTarget1 As Decimal
        Public BuyTarget2 As Decimal
        Public BuyTarget3 As Decimal
        Public BuyTarget4 As Decimal
        Public BuyTarget5 As Decimal
        Public SellLevel As Decimal
        Public SellEntry As Decimal
        Public SellStoploss As Decimal
        Public SellTarget1 As Decimal
        Public SellTarget2 As Decimal
        Public SellTarget3 As Decimal
        Public SellTarget4 As Decimal
        Public SellTarget5 As Decimal
    End Class
End Class
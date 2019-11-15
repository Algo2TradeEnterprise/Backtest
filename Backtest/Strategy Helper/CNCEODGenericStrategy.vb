﻿Imports Algo2TradeBLL
Imports System.IO
Imports System.Threading
Imports Utilities.Strings

Namespace StrategyHelper
    Public Class CNCEODGenericStrategy
        Inherits Strategy
        Implements IDisposable

        Private EODPayload As Dictionary(Of String, Dictionary(Of Date, Payload))

        Public Property StockFileName As String
        Public Property RuleNumber As Integer
        Public Property RuleEntityData As RuleEntities
        Public Sub New(ByVal canceller As CancellationTokenSource,
                       ByVal exchangeStartTime As TimeSpan,
                       ByVal exchangeEndTime As TimeSpan,
                       ByVal tradeStartTime As TimeSpan,
                       ByVal lastTradeEntryTime As TimeSpan,
                       ByVal eodExitTime As TimeSpan,
                       ByVal tickSize As Decimal,
                       ByVal marginMultiplier As Decimal,
                       ByVal timeframe As Integer,
                       ByVal heikenAshiCandle As Boolean,
                       ByVal stockType As Trade.TypeOfStock,
                       ByVal databaseTable As Common.DataBaseTable,
                       ByVal dataSource As SourceOfData,
                       ByVal initialCapital As Decimal,
                       ByVal usableCapital As Decimal,
                       ByVal minimumEarnedCapitalToWithdraw As Decimal,
                       ByVal amountToBeWithdrawn As Decimal)
            MyBase.New(canceller, exchangeStartTime, exchangeEndTime, tradeStartTime, lastTradeEntryTime, eodExitTime, tickSize, marginMultiplier, timeframe, heikenAshiCandle, stockType, Trade.TypeOfTrade.CNC, databaseTable, dataSource, initialCapital, usableCapital, minimumEarnedCapitalToWithdraw, amountToBeWithdrawn)
        End Sub

        Public Overrides Async Function TestStrategyAsync(startDate As Date, endDate As Date, filename As String) As Task
            If Not Me.ExitOnOverAllFixedTargetStoploss Then
                Me.OverAllProfitPerDay = Decimal.MaxValue
                Me.OverAllLossPerDay = Decimal.MinValue
            End If
            If Not Me.ExitOnStockFixedTargetStoploss Then
                Me.StockMaxProfitPerDay = Decimal.MaxValue
                Me.StockMaxLossPerDay = Decimal.MinValue
            End If
            If filename Is Nothing Then Throw New ApplicationException("Filename Invalid")
            Dim tradesFileName As String = Path.Combine(My.Application.Info.DirectoryPath, String.Format("{0}.Trades.a2t", filename))
            Dim capitalFileName As String = Path.Combine(My.Application.Info.DirectoryPath, String.Format("{0}.Capital.a2t", filename))

            If File.Exists(tradesFileName) AndAlso File.Exists(capitalFileName) Then
                Dim folderpath As String = Path.Combine(My.Application.Info.DirectoryPath, "BackTest Output")
                Dim files() = Directory.GetFiles(folderpath, "*.xlsx")
                For Each file In files
                    If file.ToUpper.Contains(filename.ToUpper) Then
                        Exit Function
                    End If
                Next
                PrintArrayToExcel(filename, tradesFileName, capitalFileName)
            Else
                If File.Exists(tradesFileName) Then File.Delete(tradesFileName)
                If File.Exists(capitalFileName) Then File.Delete(capitalFileName)
                Dim totalPL As Decimal = 0
                Dim tradeCheckingDate As Date = startDate.Date
                TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))
                Me.AvailableCapital = Me.UsableCapital
                While tradeCheckingDate <= endDate.Date
                    Console.WriteLine(String.Format("Capital available:{0},{1},", tradeCheckingDate.ToShortDateString, Me.AvailableCapital))
                    Dim tradingDay As Boolean = True
                    If tradingDay Then
                        _canceller.Token.ThrowIfCancellationRequested()
                        'Me.AvailableCapital = Me.UsableCapital
                        'TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))
                        Dim stockList As Dictionary(Of String, StockDetails) = GetStockData(tradeCheckingDate)

                        _canceller.Token.ThrowIfCancellationRequested()
                        If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                            OnHeartbeat("Deserializing candles")
                            If EODPayload Is Nothing Then
                                For Each stock In stockList.Keys
                                    If EODPayload Is Nothing Then EODPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                    Dim eodFilename As String = Path.Combine(My.Application.Info.DirectoryPath, "Candle Data", String.Format("{0} EOD Data.a2t", stock))
                                    If File.Exists(eodFilename) Then
                                        If Not EODPayload.ContainsKey(stock) Then
                                            Dim candleData As Dictionary(Of Date, Payload) = Nothing
                                            Using stream As New FileStream(eodFilename, FileMode.Open)
                                                Dim binaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
                                                While stream.Position <> stream.Length
                                                    Dim tempData As Dictionary(Of Date, Payload) = binaryFormatter.Deserialize(stream)
                                                    If tempData IsNot Nothing AndAlso tempData.Count > 0 Then
                                                        For Each runningData In tempData
                                                            If candleData Is Nothing Then candleData = New Dictionary(Of Date, Payload)
                                                            If Not candleData.ContainsKey(runningData.Key) Then
                                                                candleData.Add(runningData.Key, runningData.Value)
                                                            End If
                                                        Next
                                                    End If
                                                End While
                                            End Using
                                            EODPayload.Add(stock, candleData)
                                        End If
                                    End If
                                Next
                            End If

                            Dim currentDayStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                            Dim XDayStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                            Dim stocksRuleData As Dictionary(Of String, StrategyRule) = Nothing

                            'First lets build the payload for all the stocks
                            Dim stockCount As Integer = 0
                            Dim eligibleStockCount As Integer = 0
                            For Each stock In stockList.Keys
                                _canceller.Token.ThrowIfCancellationRequested()
                                stockCount += 1
                                Dim XDayPayload As Dictionary(Of Date, Payload) = Nothing
                                Dim currentDayPayload As Dictionary(Of Date, Payload) = Nothing
                                If Me.DataSource = SourceOfData.Database Then
                                    XDayPayload = Await GetCandleData(Me.DatabaseTable, stock, tradeCheckingDate.AddYears(-3), tradeCheckingDate).ConfigureAwait(False)
                                ElseIf Me.DataSource = SourceOfData.Live Then
                                    XDayPayload = Await GetCandleData(Me.DatabaseTable, stock, tradeCheckingDate.AddYears(-3), tradeCheckingDate).ConfigureAwait(False)
                                End If

                                _canceller.Token.ThrowIfCancellationRequested()
                                'Now transfer only the current date payload into the workable payload (this will be used for the main loop and checking if the date is a valid date)
                                If XDayPayload IsNot Nothing AndAlso XDayPayload.Count > 0 Then
                                    OnHeartbeat(String.Format("Processing for {0} on {1}. Stock Counter: [ {2}/{3} ]", stock, tradeCheckingDate.ToShortDateString, stockCount, stockList.Count))

                                    If XDayStocksPayload Is Nothing Then XDayStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                    XDayStocksPayload.Add(stock, XDayPayload)

                                    For Each runningPayload In XDayPayload.Keys
                                        _canceller.Token.ThrowIfCancellationRequested()
                                        If runningPayload.Date = tradeCheckingDate.Date Then
                                            If currentDayPayload Is Nothing Then currentDayPayload = New Dictionary(Of Date, Payload)
                                            currentDayPayload.Add(runningPayload, XDayPayload(runningPayload))
                                            Exit For
                                        End If
                                    Next
                                    'Add all these payloads into the stock collections
                                    If currentDayPayload IsNot Nothing AndAlso currentDayPayload.Count > 0 Then
                                        If currentDayStocksPayload Is Nothing Then currentDayStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                        currentDayStocksPayload.Add(stock, currentDayPayload)
                                        If stocksRuleData Is Nothing Then stocksRuleData = New Dictionary(Of String, StrategyRule)
                                        Dim stockRule As StrategyRule = Nothing

                                        Dim tradingSymbol As String = currentDayPayload.LastOrDefault.Value.TradingSymbol
                                        Select Case RuleNumber
                                            Case 0
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 1
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 2
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 3
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 4
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 5
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 6
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 7
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 8
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 9
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 10
                                                stockRule = New VijayCNCStrategyRule(XDayPayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                            Case 11
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 12
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 13
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 14
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 15
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 16
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 17
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 18
                                                stockRule = New InvestmentCNCStrategyRule(XDayPayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData, stockList(stock).Supporting1)
                                            Case 19
                                                Throw New ApplicationException("Not a CNC strategy")
                                            Case 20
                                                Throw New ApplicationException("Not a CNC strategy")
                                        End Select

                                        AddHandler stockRule.Heartbeat, AddressOf OnHeartbeat
                                        stockRule.CompletePreProcessing()
                                        stocksRuleData.Add(stock, stockRule)
                                    End If
                                End If
                            Next
                            '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                            If currentDayStocksPayload IsNot Nothing AndAlso currentDayStocksPayload.Count > 0 Then
                                OnHeartbeat(String.Format("Checking Trade on {0}", tradeCheckingDate.ToShortDateString))
                                _canceller.Token.ThrowIfCancellationRequested()
                                For Each stockName In stockList
                                    stockName.Value.PlaceOrderDoneForTheMinute = False
                                    stockName.Value.ExitOrderDoneForTheMinute = False
                                    stockName.Value.CancelOrderDoneForTheMinute = False
                                    stockName.Value.ModifyStoplossOrderDoneForTheMinute = False
                                    stockName.Value.ModifyTargetOrderDoneForTheMinute = False
                                Next
                                For Each stockName In stockList.Keys
                                    _canceller.Token.ThrowIfCancellationRequested()

                                    If Not stocksRuleData.ContainsKey(stockName) Then
                                        Continue For
                                    End If
                                    Dim stockStrategyRule As StrategyRule = stocksRuleData(stockName)

                                    If Not stockList(stockName).EligibleToTakeTrade Then
                                        Continue For
                                    End If

                                    If Not currentDayStocksPayload.ContainsKey(stockName) Then
                                        Continue For
                                    End If

                                    'Get the current candle from the stock collection for this stock for that day
                                    _canceller.Token.ThrowIfCancellationRequested()
                                    Dim currentDayCandlePayload As Payload = Nothing
                                    If currentDayStocksPayload.ContainsKey(stockName) AndAlso
                                        currentDayStocksPayload(stockName).ContainsKey(tradeCheckingDate.Date) Then
                                        currentDayCandlePayload = currentDayStocksPayload(stockName)(tradeCheckingDate.Date)
                                    End If

                                    'Now check trade
                                    _canceller.Token.ThrowIfCancellationRequested()
                                    If currentDayCandlePayload IsNot Nothing Then
                                        Dim runningTick As Payload = New Payload(Payload.CandleDataSource.Calculated) With {
                                            .Open = currentDayCandlePayload.Close,
                                            .Low = currentDayCandlePayload.Close,
                                            .High = currentDayCandlePayload.Close,
                                            .Close = currentDayCandlePayload.Close,
                                            .PayloadDate = currentDayCandlePayload.PayloadDate,
                                            .TradingSymbol = currentDayCandlePayload.TradingSymbol
                                        }

                                        SetCurrentLTPForStock(currentDayCandlePayload, runningTick, Trade.TypeOfTrade.CNC)

                                        'Update Collection
                                        Await stockStrategyRule.UpdateRequiredCollectionsAsync(runningTick).ConfigureAwait(False)

                                        'Exit Trade From Rule
                                        _canceller.Token.ThrowIfCancellationRequested()
                                        Dim potentialRuleCancelTrades As List(Of Trade) = GetSpecificTrades(currentDayCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Open)
                                        If potentialRuleCancelTrades IsNot Nothing AndAlso potentialRuleCancelTrades.Count > 0 Then
                                            If Me.TickBasedStrategy OrElse Not stockList(stockName).CancelOrderDoneForTheMinute Then
                                                For Each runningCancelTrade In potentialRuleCancelTrades
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim exitOrderDetails As Tuple(Of Boolean, Decimal, String) = Await stockStrategyRule.IsTriggerReceivedForExitCNCEODOrderAsync(runningTick, runningCancelTrade).ConfigureAwait(False)
                                                    If exitOrderDetails IsNot Nothing AndAlso exitOrderDetails.Item1 Then
                                                        _canceller.Token.ThrowIfCancellationRequested()
                                                        Dim exitTick As Payload = New Payload(Payload.CandleDataSource.Calculated) With {
                                                                    .Open = exitOrderDetails.Item2,
                                                                    .Low = exitOrderDetails.Item2,
                                                                    .High = exitOrderDetails.Item2,
                                                                    .Close = exitOrderDetails.Item2,
                                                                    .PayloadDate = runningTick.PayloadDate,
                                                                    .TradingSymbol = runningTick.TradingSymbol
                                                                }
                                                        ExitTradeByForce(runningCancelTrade, exitTick, exitOrderDetails.Item2)
                                                    End If
                                                Next
                                                stockList(stockName).CancelOrderDoneForTheMinute = True
                                            End If
                                        End If
                                        _canceller.Token.ThrowIfCancellationRequested()
                                        Dim potentialRuleExitTrades As List(Of Trade) = GetSpecificTrades(currentDayCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
                                        If potentialRuleExitTrades IsNot Nothing AndAlso potentialRuleExitTrades.Count > 0 Then
                                            If Me.TickBasedStrategy OrElse Not stockList(stockName).ExitOrderDoneForTheMinute Then
                                                For Each runningExitTrade In potentialRuleExitTrades
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim exitOrderDetails As Tuple(Of Boolean, Decimal, String) = Await stockStrategyRule.IsTriggerReceivedForExitCNCEODOrderAsync(runningTick, runningExitTrade).ConfigureAwait(False)
                                                    If exitOrderDetails IsNot Nothing AndAlso exitOrderDetails.Item1 Then
                                                        _canceller.Token.ThrowIfCancellationRequested()
                                                        Dim exitTick As Payload = New Payload(Payload.CandleDataSource.Calculated) With {
                                                                    .Open = exitOrderDetails.Item2,
                                                                    .Low = exitOrderDetails.Item2,
                                                                    .High = exitOrderDetails.Item2,
                                                                    .Close = exitOrderDetails.Item2,
                                                                    .PayloadDate = runningTick.PayloadDate,
                                                                    .TradingSymbol = runningTick.TradingSymbol
                                                                }
                                                        ExitTradeByForce(runningExitTrade, exitTick, exitOrderDetails.Item3)
                                                    End If
                                                Next
                                                stockList(stockName).ExitOrderDoneForTheMinute = True
                                            End If
                                        End If

                                        'Place Order
                                        _canceller.Token.ThrowIfCancellationRequested()
                                        Dim placeOrderDetails As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
                                        If Me.StockNumberOfTrades(runningTick.PayloadDate, runningTick.TradingSymbol) < Me.NumberOfTradesPerStockPerDay AndAlso
                                        Me.TotalPLAfterBrokerage(runningTick.PayloadDate) < Me.OverAllProfitPerDay AndAlso
                                        Me.TotalPLAfterBrokerage(runningTick.PayloadDate) > Math.Abs(Me.OverAllLossPerDay) * -1 AndAlso
                                        Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol) < Me.StockMaxProfitPerDay AndAlso
                                        Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol) > Math.Abs(Me.StockMaxLossPerDay) * -1 AndAlso
                                        Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol) < stockStrategyRule.MaxProfitOfThisStock AndAlso
                                        Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol) > Math.Abs(stockStrategyRule.MaxLossOfThisStock) * -1 Then
                                            placeOrderDetails = Await stockStrategyRule.IsTriggerReceivedForPlaceOrderAsync(runningTick).ConfigureAwait(False)
                                            stockList(stockName).PlaceOrderTrigger = placeOrderDetails
                                            stockList(stockName).PlaceOrderDoneForTheMinute = True
                                        End If
                                        If placeOrderDetails IsNot Nothing AndAlso placeOrderDetails.Item1 Then
                                            Dim placeOrders As List(Of PlaceOrderParameters) = placeOrderDetails.Item2
                                            If placeOrders IsNot Nothing AndAlso placeOrders.Count > 0 Then
                                                Dim tradeTag As String = System.Guid.NewGuid.ToString()
                                                For Each runningOrder In placeOrders
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    If runningOrder.Used Then
                                                        Continue For
                                                    End If
                                                    Dim runningTrade As Trade = New Trade(originatingStrategy:=Me,
                                                                                          tradingSymbol:=runningTick.TradingSymbol,
                                                                                          stockType:=Me.StockType,
                                                                                          orderType:=runningOrder.OrderType,
                                                                                          tradingDate:=runningTick.PayloadDate,
                                                                                          entryDirection:=runningOrder.EntryDirection,
                                                                                          entryPrice:=runningOrder.EntryPrice,
                                                                                          entryBuffer:=runningOrder.Buffer,
                                                                                          squareOffType:=Trade.TypeOfTrade.CNC,
                                                                                          entryCondition:=Trade.TradeEntryCondition.Original,
                                                                                          entryRemark:="Original Entry",
                                                                                          quantity:=runningOrder.Quantity,
                                                                                          lotSize:=stockStrategyRule.LotSize,
                                                                                          potentialTarget:=runningOrder.Target,
                                                                                          targetRemark:=Math.Abs(runningOrder.EntryPrice - runningOrder.Target),
                                                                                          potentialStopLoss:=runningOrder.Stoploss,
                                                                                          stoplossBuffer:=runningOrder.Buffer,
                                                                                          slRemark:=Math.Abs(runningOrder.EntryPrice - runningOrder.Stoploss),
                                                                                          signalCandle:=runningOrder.SignalCandle)

                                                    runningTrade.UpdateTrade(Tag:=tradeTag,
                                                                             SquareOffValue:=Math.Abs(runningOrder.EntryPrice - runningOrder.Target),
                                                                             Supporting1:=runningOrder.Supporting1,
                                                                             Supporting2:=runningOrder.Supporting2,
                                                                             Supporting3:=runningOrder.Supporting3,
                                                                             Supporting4:=runningOrder.Supporting4,
                                                                             Supporting5:=runningOrder.Supporting5)

                                                    If PlaceOrModifyOrder(runningTrade, Nothing) Then
                                                        runningOrder.Used = True
                                                    End If
                                                Next
                                            End If
                                        End If

                                        'Enter Trade
                                        _canceller.Token.ThrowIfCancellationRequested()
                                        Dim potentialEntryTrades As List(Of Trade) = GetSpecificTrades(currentDayCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Open)
                                        If potentialEntryTrades IsNot Nothing AndAlso potentialEntryTrades.Count > 0 Then
                                            For Each runningPotentialEntryTrade In potentialEntryTrades
                                                _canceller.Token.ThrowIfCancellationRequested()
                                                Dim entryTick As Payload = New Payload(Payload.CandleDataSource.Calculated) With {
                                                                    .Open = runningPotentialEntryTrade.EntryPrice,
                                                                    .Low = runningPotentialEntryTrade.EntryPrice,
                                                                    .High = runningPotentialEntryTrade.EntryPrice,
                                                                    .Close = runningPotentialEntryTrade.EntryPrice,
                                                                    .PayloadDate = runningTick.PayloadDate,
                                                                    .TradingSymbol = runningTick.TradingSymbol
                                                                }
                                                EnterTradeIfPossible(runningPotentialEntryTrade, entryTick)
                                            Next
                                        End If

                                        'Modify Stoploss Trade
                                        _canceller.Token.ThrowIfCancellationRequested()
                                        Dim potentialModifySLTrades As List(Of Trade) = GetSpecificTrades(currentDayCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
                                        If potentialModifySLTrades IsNot Nothing AndAlso potentialModifySLTrades.Count > 0 Then
                                            If Me.TickBasedStrategy OrElse Not stockList(stockName).ModifyStoplossOrderDoneForTheMinute OrElse Me.TrailingStoploss Then
                                                For Each runningModifyTrade In potentialModifySLTrades
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim modifyOrderDetails As Tuple(Of Boolean, Decimal, String) = Await stockStrategyRule.IsTriggerReceivedForModifyStoplossOrderAsync(runningTick, runningModifyTrade).ConfigureAwait(False)
                                                    If modifyOrderDetails IsNot Nothing AndAlso modifyOrderDetails.Item1 Then
                                                        _canceller.Token.ThrowIfCancellationRequested()
                                                        runningModifyTrade.UpdateTrade(PotentialStopLoss:=modifyOrderDetails.Item2, SLRemark:=modifyOrderDetails.Item3)
                                                    End If
                                                Next
                                                stockList(stockName).ModifyStoplossOrderDoneForTheMinute = True
                                            End If
                                        End If

                                        'Modify Target Trade
                                        _canceller.Token.ThrowIfCancellationRequested()
                                        Dim potentialModifyTargetTrades As List(Of Trade) = GetSpecificTrades(currentDayCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
                                        If potentialModifyTargetTrades IsNot Nothing AndAlso potentialModifyTargetTrades.Count > 0 Then
                                            If Me.TickBasedStrategy OrElse Not stockList(stockName).ModifyTargetOrderDoneForTheMinute Then
                                                For Each runningModifyTrade In potentialModifyTargetTrades
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim modifyOrderDetails As Tuple(Of Boolean, Decimal, String) = Await stockStrategyRule.IsTriggerReceivedForModifyTargetOrderAsync(runningTick, runningModifyTrade).ConfigureAwait(False)
                                                    If modifyOrderDetails IsNot Nothing AndAlso modifyOrderDetails.Item1 Then
                                                        _canceller.Token.ThrowIfCancellationRequested()
                                                        runningModifyTrade.UpdateTrade(PotentialTarget:=modifyOrderDetails.Item2, TargetRemark:=modifyOrderDetails.Item3)
                                                    End If
                                                Next
                                                stockList(stockName).ModifyTargetOrderDoneForTheMinute = True
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                            If tradeCheckingDate.Date = endDate.Date Then
                                If XDayStocksPayload IsNot Nothing AndAlso XDayStocksPayload.Count > 0 Then
                                    Dim currentDayStocksExitPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                                    For Each runningStock In XDayStocksPayload.Keys
                                        Dim lastAvailableDayPayload As KeyValuePair(Of Date, Payload) =
                                            XDayStocksPayload(runningStock).Where(Function(X)
                                                                                      Return X.Key.Date <= tradeCheckingDate.Date
                                                                                  End Function).LastOrDefault
                                        If currentDayStocksExitPayload Is Nothing Then currentDayStocksExitPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                        Dim dummyPayload As Payload = New Payload(Payload.CandleDataSource.Calculated) With {
                                            .Open = lastAvailableDayPayload.Value.Close,
                                            .Low = lastAvailableDayPayload.Value.Close,
                                            .High = lastAvailableDayPayload.Value.Close,
                                            .Close = lastAvailableDayPayload.Value.Close,
                                            .PayloadDate = lastAvailableDayPayload.Value.PayloadDate,
                                            .TradingSymbol = lastAvailableDayPayload.Value.TradingSymbol
                                        }
                                        currentDayStocksExitPayload.Add(runningStock, New Dictionary(Of Date, Payload) From {{dummyPayload.PayloadDate, dummyPayload}})
                                    Next
                                    ExitAllTradeByForce(tradeCheckingDate.AddDays(1).Date, currentDayStocksExitPayload, Trade.TypeOfTrade.CNC, "Open Trade")
                                End If
                            End If
                        End If
                        SetOverallDrawUpDrawDownForTheDay(tradeCheckingDate)
                        totalPL += TotalPLAfterBrokerage(tradeCheckingDate)
                    End If
                    tradeCheckingDate = tradeCheckingDate.AddDays(1)
                End While   'Date Loop

                'Serialization
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                    OnHeartbeat("Serializing Trades collection")
                    SerializeFromCollectionUsingFileStream(Of Dictionary(Of Date, Dictionary(Of String, List(Of Trade))))(tradesFileName, TradesTaken, False)
                End If
                If CapitalMovement IsNot Nothing Then Utilities.Strings.SerializeFromCollection(Of Dictionary(Of Date, List(Of Capital)))(capitalFileName, CapitalMovement)

                PrintArrayToExcel(filename, tradesFileName, capitalFileName)
            End If
        End Function

#Region "Private Function"
        Private Async Function GetCandleData(ByVal tablename As Common.DataBaseTable, ByVal rawInstrumentName As String,
                                             ByVal startDate As Date, ByVal endDate As Date) As Task(Of Dictionary(Of Date, Payload))
            OnHeartbeat(String.Format("Getting Candle Data for {0} on {1}", rawInstrumentName, endDate.ToString("yyyy-MM-dd")))
            Dim ret As Dictionary(Of Date, Payload) = Nothing
            Dim fileName As String = Path.Combine(My.Application.Info.DirectoryPath, "Candle Data", String.Format("{0} EOD Data.a2t", rawInstrumentName))
            Dim dataNotAvailable As Boolean = False
            If File.Exists(fileName) Then
                If EODPayload IsNot Nothing AndAlso EODPayload.ContainsKey(rawInstrumentName) Then
                    ret = EODPayload(rawInstrumentName)
                    If Not ret.ContainsKey(endDate.Date) Then dataNotAvailable = True
                Else
                    dataNotAvailable = True
                End If
            Else
                dataNotAvailable = True
            End If

            If dataNotAvailable Then
                If Not endDate.DayOfWeek = DayOfWeek.Sunday AndAlso Not endDate.DayOfWeek = DayOfWeek.Saturday Then
                    If Me.DataSource = SourceOfData.Database Then
                        ret = Cmn.GetRawPayload(tablename, rawInstrumentName, startDate, endDate)
                    ElseIf Me.DataSource = SourceOfData.Live Then
                        ret = Await Cmn.GetHistoricalDataAsync(tablename, rawInstrumentName, startDate, endDate).ConfigureAwait(False)
                    End If
                End If
                If ret IsNot Nothing AndAlso ret.Count > 0 Then
                    Dim dataToSerialise As Dictionary(Of Date, Payload) = Nothing
                    If File.Exists(fileName) Then
                        If EODPayload IsNot Nothing AndAlso EODPayload.ContainsKey(rawInstrumentName) Then
                            Dim availableData As Dictionary(Of Date, Payload) = EODPayload(rawInstrumentName)
                            For Each runningPayload In ret
                                If Not availableData.ContainsKey(runningPayload.Key) Then
                                    If dataToSerialise Is Nothing Then dataToSerialise = New Dictionary(Of Date, Payload)
                                    dataToSerialise.Add(runningPayload.Key, runningPayload.Value)
                                End If
                            Next
                        Else
                            dataToSerialise = ret
                        End If
                    Else
                        dataToSerialise = ret
                    End If
                    If dataToSerialise IsNot Nothing AndAlso dataToSerialise.Count > 0 Then
                        SerializeFromCollectionUsingFileStream(Of Dictionary(Of Date, Payload))(fileName, dataToSerialise)
                    End If
                End If
            End If
            Return ret
        End Function
#End Region

#Region "Stock Selection"
        Private Function GetStockData(tradingDate As Date) As Dictionary(Of String, StockDetails)
            Dim ret As Dictionary(Of String, StockDetails) = Nothing
            If Me.StockFileName IsNot Nothing Then
                Dim dt As DataTable = Nothing
                Using csvHelper As New Utilities.DAL.CSVHelper(Me.StockFileName, ",", _canceller)
                    dt = csvHelper.GetDataTableFromCSV(1)
                End Using
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Select Case Me.RuleNumber
                        Case 10
                            Dim counter As Integer = 0
                            For i = 1 To dt.Rows.Count - 1
                                Dim rowDate As Date = dt.Rows(i)(0)
                                'If rowDate.Date = tradingDate.Date Then
                                If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                                Dim tradingSymbol As String = dt.Rows(i).Item(1)
                                Dim instrumentName As String = Nothing
                                If tradingSymbol.Contains("FUT") Then
                                    instrumentName = tradingSymbol.Remove(tradingSymbol.Count - 8)
                                Else
                                    instrumentName = tradingSymbol
                                End If
                                Dim detailsOfStock As StockDetails = New StockDetails With
                                    {.StockName = instrumentName,
                                    .LotSize = dt.Rows(i).Item(2),
                                    .EligibleToTakeTrade = True,
                                    .Supporting1 = dt.Rows(i).Item(5)}
                                ret.Add(instrumentName, detailsOfStock)
                                counter += 1
                                If counter = Me.NumberOfTradeableStockPerDay Then Exit For
                                'End If
                            Next
                        Case 18
                            For i = 1 To dt.Rows.Count - 1
                                Dim rowDate As Date = dt.Rows(i)(0)
                                'If rowDate.Date = tradingDate.Date Then
                                If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                                Dim tradingSymbol As String = dt.Rows(i).Item(1)
                                Dim instrumentName As String = Nothing
                                If tradingSymbol.Contains("FUT") Then
                                    instrumentName = tradingSymbol.Remove(tradingSymbol.Count - 8)
                                Else
                                    instrumentName = tradingSymbol
                                End If
                                Dim detailsOfStock As StockDetails = New StockDetails With
                                    {.StockName = instrumentName,
                                    .LotSize = 1,
                                    .EligibleToTakeTrade = True,
                                    .Supporting1 = dt.Rows(i).Item(2)}
                                ret.Add(instrumentName, detailsOfStock)
                                If i = Me.NumberOfTradeableStockPerDay Then Exit For
                                'End If
                            Next
                    End Select
                End If
            End If
            Return ret
        End Function
#End Region

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            ' TODO: uncomment the following line if Finalize() is overridden above.
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class
End Namespace
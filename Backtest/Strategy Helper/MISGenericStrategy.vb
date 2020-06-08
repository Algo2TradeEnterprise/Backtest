Imports Algo2TradeBLL
Imports System.IO
Imports System.Threading
Imports Utilities.Strings

Namespace StrategyHelper
    Public Class MISGenericStrategy
        Inherits Strategy
        Implements IDisposable
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
                       ByVal optionStockType As Trade.TypeOfStock,
                       ByVal databaseTable As Common.DataBaseTable,
                       ByVal dataSource As SourceOfData,
                       ByVal initialCapital As Decimal,
                       ByVal usableCapital As Decimal,
                       ByVal minimumEarnedCapitalToWithdraw As Decimal,
                       ByVal amountToBeWithdrawn As Decimal)
            MyBase.New(canceller, exchangeStartTime, exchangeEndTime, tradeStartTime, lastTradeEntryTime, eodExitTime, tickSize, marginMultiplier, timeframe, heikenAshiCandle, stockType, optionStockType, Trade.TypeOfTrade.MIS, databaseTable, dataSource, initialCapital, usableCapital, minimumEarnedCapitalToWithdraw, amountToBeWithdrawn)
            Me.NumberOfTradesPerDay = Integer.MaxValue
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

            If filename Is Nothing Then Throw New ApplicationException("Invalid Filename")
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
                Dim strategyName As String = String.Format("Strategy{0}", Me.RuleNumber)
                OnHeartbeat("Getting unique instrument list")
                Dim allInstrumentList As List(Of String) = GetUniqueInstrumentList(startDate, endDate)
                Dim dataFtchr As DataFetcher = New DataFetcher(_canceller, My.Settings.ServerName, allInstrumentList, startDate.AddDays(-15), endDate, Me.StockType, Me.OptionStockType, strategyName)
                AddHandler dataFtchr.Heartbeat, AddressOf OnHeartbeat
                AddHandler dataFtchr.WaitingFor, AddressOf OnWaitingFor
                AddHandler dataFtchr.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
                AddHandler dataFtchr.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete

                If File.Exists(tradesFileName) Then File.Delete(tradesFileName)
                If File.Exists(capitalFileName) Then File.Delete(capitalFileName)
                Dim totalPL As Decimal = 0
                Dim tradeCheckingDate As Date = startDate.Date
                Dim portfolioLossPerDay As Decimal = Me.OverAllLossPerDay
                While tradeCheckingDate <= endDate.Date
                    _canceller.Token.ThrowIfCancellationRequested()
                    Me.AvailableCapital = Me.UsableCapital
                    Me.OverAllLossPerDay = portfolioLossPerDay
                    TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))
                    Dim stockList As Dictionary(Of String, StockDetails) = Await GetStockDataAsync(tradeCheckingDate).ConfigureAwait(False)

                    _canceller.Token.ThrowIfCancellationRequested()
                    If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                        Dim currentDayOneMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                        Dim XDayOneMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                        Dim stocksRuleData As Dictionary(Of String, StrategyRule) = Nothing

                        'First lets build the payload for all the stocks
                        Dim stockCount As Integer = 0
                        Dim eligibleStockCount As Integer = 0
                        For Each stock In stockList.Keys
                            _canceller.Token.ThrowIfCancellationRequested()
                            stockCount += 1
                            OnHeartbeat(String.Format("Getting data for {0} on {1}. Stock Counter: [ {2}/{3} ]", stock, tradeCheckingDate.ToShortDateString, stockCount, stockList.Count))
                            Dim XDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                            Dim currentDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                            If Me.DataSource = SourceOfData.Database Then
                                'XDayOneMinutePayload = Cmn.GetRawPayload(Me.DatabaseTable, stock, tradeCheckingDate.AddDays(-7), tradeCheckingDate)
                                XDayOneMinutePayload = Await dataFtchr.GetCandleData(stockList(stock).TradingSymbol, tradeCheckingDate.AddDays(-7), tradeCheckingDate).ConfigureAwait(False)
                            ElseIf Me.DataSource = SourceOfData.Live Then
                                XDayOneMinutePayload = Await Cmn.GetHistoricalDataAsync(Me.DatabaseTable, stock, tradeCheckingDate.AddDays(-7), tradeCheckingDate).ConfigureAwait(False)
                            End If

                            _canceller.Token.ThrowIfCancellationRequested()
                            'Now transfer only the current date payload into the workable payload (this will be used for the main loop and checking if the date is a valid date)
                            If XDayOneMinutePayload IsNot Nothing AndAlso XDayOneMinutePayload.Count > 0 Then
                                OnHeartbeat(String.Format("Processing for {0} on {1}. Stock Counter: [ {2}/{3} ]", stock, tradeCheckingDate.ToShortDateString, stockCount, stockList.Count))
                                For Each runningPayload In XDayOneMinutePayload.Keys
                                    _canceller.Token.ThrowIfCancellationRequested()
                                    If runningPayload.Date = tradeCheckingDate.Date Then
                                        If currentDayOneMinutePayload Is Nothing Then currentDayOneMinutePayload = New Dictionary(Of Date, Payload)
                                        currentDayOneMinutePayload.Add(runningPayload, XDayOneMinutePayload(runningPayload))
                                    End If
                                Next
                                'Add all these payloads into the stock collections
                                If currentDayOneMinutePayload IsNot Nothing AndAlso currentDayOneMinutePayload.Count > 0 Then
                                    If currentDayOneMinuteStocksPayload Is Nothing Then currentDayOneMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                    currentDayOneMinuteStocksPayload.Add(stock, currentDayOneMinutePayload)
                                    If XDayOneMinuteStocksPayload Is Nothing Then XDayOneMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                    XDayOneMinuteStocksPayload.Add(stock, XDayOneMinutePayload)
                                    Dim stockRule As StrategyRule = Nothing

                                    Dim tradingSymbol As String = currentDayOneMinutePayload.LastOrDefault.Value.TradingSymbol
                                    Select Case RuleNumber
                                        Case 0
                                            stockRule = New ReversalHHLLBreakoutStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 1
                                            stockRule = New FractalTrendLineStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 2
                                            stockRule = New MarketPlusMarketMinusStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData, stockList(stock).Supporting1, stockList(stock).Supporting2)
                                        Case 3
                                            stockRule = New HighestLowestPointStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 4
                                            stockRule = New HeikenashiReverseSlabStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData, stockList(stock).Slab)
                                        Case 5
                                            stockRule = New EMAScalpingStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 6
                                            stockRule = New SupertrendCutReversalStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 7
                                            stockRule = New BNFMartingaleStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 8
                                            stockRule = New HL_LHBreakoutStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 9
                                            stockRule = New AlwaysInTradeMartingaleStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData, stockList(stock).Slab)
                                        Case 10
                                            stockRule = New MartingaleStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 11
                                            stockRule = New AnchorSatelliteHKStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 12
                                            stockRule = New SmallOpeningRangeBreakoutStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 13
                                            stockRule = New LossMakeupFavourableFractalBreakoutStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 14
                                            stockRule = New HKReverseSlabMartingaleStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 15
                                            stockRule = New LowerPriceOptionBuyOnlyStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData, stockList(stock).Supporting1, stockList(stock).Supporting2)
                                        Case 16
                                            stockRule = New LowerPriceOptionOIChangeBuyOnlyStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 17
                                            stockRule = New LowerPriceOptionBuyOnlyEODStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData, stockList(stock).Supporting1)
                                        Case 18
                                            stockRule = New EveryMinuteTopGainerLosserHKReversalStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData, stockList(stock).SupportingDate, stockList(stock).Supporting1)
                                        Case 19
                                            stockRule = New LossMakeupRainbowStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 20
                                            stockRule = New AnchorSatelliteLossMakeupStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                        Case 21
                                            stockRule = New AnchorSatelliteLossMakeupHKStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                    End Select

                                    AddHandler stockRule.Heartbeat, AddressOf OnHeartbeat
                                    stockRule.CompletePreProcessing()
                                    If stocksRuleData Is Nothing Then stocksRuleData = New Dictionary(Of String, StrategyRule)
                                    stocksRuleData.Add(stock, stockRule)
                                End If
                            End If
                        Next
                        '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                        If currentDayOneMinuteStocksPayload IsNot Nothing AndAlso currentDayOneMinuteStocksPayload.Count > 0 Then
                            OnHeartbeat(String.Format("Checking Trade on {0}", tradeCheckingDate.ToShortDateString))
                            _canceller.Token.ThrowIfCancellationRequested()
                            Dim eodTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, Me.EODExitTime.Hours, Me.EODExitTime.Minutes, Me.EODExitTime.Seconds)
                            Dim startMinute As TimeSpan = Me.ExchangeStartTime
                            Dim endMinute As TimeSpan = ExchangeEndTime
                            While startMinute < endMinute
                                _canceller.Token.ThrowIfCancellationRequested()
                                OnHeartbeat(String.Format("Checking Trade on {0}. Time:{1}", tradeCheckingDate.ToShortDateString, startMinute.ToString))
                                Dim startSecond As TimeSpan = startMinute
                                Dim endSecond As TimeSpan = startMinute.Add(TimeSpan.FromMinutes(Me.SignalTimeFrame - 1))
                                endSecond = endSecond.Add(TimeSpan.FromSeconds(59))
                                Dim potentialCandleSignalTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startMinute.Hours, startMinute.Minutes, startMinute.Seconds)
                                Dim potentialTickSignalTime As Date = Nothing

                                _canceller.Token.ThrowIfCancellationRequested()
                                While startSecond <= endSecond
                                    potentialTickSignalTime = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startSecond.Hours, startSecond.Minutes, startSecond.Seconds)
                                    If potentialTickSignalTime.Second = 0 Then
                                        potentialCandleSignalTime = potentialTickSignalTime
                                        For Each stockName In stockList
                                            stockName.Value.PlaceOrderDoneForTheMinute = False
                                            stockName.Value.ExitOrderDoneForTheMinute = False
                                            stockName.Value.CancelOrderDoneForTheMinute = False
                                            stockName.Value.ModifyStoplossOrderDoneForTheMinute = False
                                            stockName.Value.ModifyTargetOrderDoneForTheMinute = False
                                        Next
                                    End If
                                    For Each stockName In stockList.Keys
                                        _canceller.Token.ThrowIfCancellationRequested()

                                        If Not stocksRuleData.ContainsKey(stockName) Then
                                            Continue For
                                        End If
                                        Dim stockStrategyRule As StrategyRule = stocksRuleData(stockName)

                                        If Not stockList(stockName).EligibleToTakeTrade Then
                                            Continue For
                                        End If

                                        If Not currentDayOneMinuteStocksPayload.ContainsKey(stockName) Then
                                            Continue For
                                        End If

                                        'Get the current minute candle from the stock collection for this stock for that day
                                        _canceller.Token.ThrowIfCancellationRequested()
                                        Dim currentMinuteCandlePayload As Payload = Nothing
                                        If currentDayOneMinuteStocksPayload.ContainsKey(stockName) AndAlso
                                            currentDayOneMinuteStocksPayload(stockName).ContainsKey(potentialCandleSignalTime) Then
                                            currentMinuteCandlePayload = currentDayOneMinuteStocksPayload(stockName)(potentialCandleSignalTime)
                                        End If

                                        'Now get the ticks for this minute and second
                                        _canceller.Token.ThrowIfCancellationRequested()
                                        Dim currentSecondTickPayload As List(Of Payload) = Nothing
                                        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.Ticks IsNot Nothing Then
                                            currentSecondTickPayload = currentMinuteCandlePayload.Ticks.FindAll(Function(x)
                                                                                                                    Return x.PayloadDate = potentialTickSignalTime
                                                                                                                End Function)

                                            _canceller.Token.ThrowIfCancellationRequested()
                                            If currentSecondTickPayload IsNot Nothing AndAlso currentSecondTickPayload.Count > 0 Then
                                                For Each runningTick In currentSecondTickPayload
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    SetCurrentLTPForStock(currentMinuteCandlePayload, runningTick, Trade.TypeOfTrade.MIS)

                                                    'Update Collection
                                                    Await stockStrategyRule.UpdateRequiredCollectionsAsync(runningTick).ConfigureAwait(False)

                                                    'Set Overall MTM
                                                    If Me.TypeOfMTMTrailing = MTMTrailingType.FixedSlabTrailing Then
                                                        Me.ExitOnOverAllFixedTargetStoploss = True
                                                        Dim trailingMTMLoss As Decimal = CalculateTrailingMTM(Me.MTMSlab, Me.MovementSlab, TotalPLAfterBrokerage(tradeCheckingDate))
                                                        If trailingMTMLoss <> Decimal.MinValue AndAlso trailingMTMLoss > Me.OverAllLossPerDay Then
                                                            Me.OverAllLossPerDay = trailingMTMLoss
                                                        End If
                                                    ElseIf Me.TypeOfMTMTrailing = MTMTrailingType.LogSlabTrailing Then
                                                        Me.ExitOnOverAllFixedTargetStoploss = True
                                                        Dim trailingMTMLoss As Decimal = CalculateLogTrailingMTM(Me.MTMSlab, TotalPLAfterBrokerage(tradeCheckingDate))
                                                        If trailingMTMLoss <> Decimal.MinValue AndAlso trailingMTMLoss > Me.OverAllLossPerDay Then
                                                            Me.OverAllLossPerDay = trailingMTMLoss
                                                        End If
                                                    ElseIf Me.TypeOfMTMTrailing = MTMTrailingType.RealtimeTrailing Then
                                                        If portfolioLossPerDay <> Decimal.MinValue AndAlso
                                                            TotalPLAfterBrokerage(tradeCheckingDate) >= Math.Abs(portfolioLossPerDay) Then
                                                            Me.ExitOnOverAllFixedTargetStoploss = True
                                                            Dim trailingMTMLoss As Decimal = TotalPLAfterBrokerage(tradeCheckingDate) * Me.RealtimeTrailingPercentage / 100
                                                            If trailingMTMLoss <> Decimal.MinValue AndAlso trailingMTMLoss > Me.OverAllLossPerDay Then
                                                                Me.OverAllLossPerDay = trailingMTMLoss
                                                            End If
                                                        End If
                                                    End If

                                                    'Specific Stock MTM Check
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim stockPL As Decimal = StockPLAfterBrokerage(tradeCheckingDate, runningTick.TradingSymbol)
                                                    If stockPL >= stockStrategyRule.MaxProfitOfThisStock Then
                                                        ExitStockTradesByForce(runningTick, Trade.TypeOfTrade.MIS, "Max Stock Profit reached for the day")
                                                        stockList(stockName).EligibleToTakeTrade = False
                                                    ElseIf stockPL <= stockStrategyRule.MaxLossOfThisStock Then
                                                        ExitStockTradesByForce(runningTick, Trade.TypeOfTrade.MIS, "Max Stock Loss reached for the day")
                                                        stockList(stockName).EligibleToTakeTrade = False
                                                    End If

                                                    'Force exit at day end
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    If runningTick.PayloadDate >= eodTime Then
                                                        ExitStockTradesByForce(runningTick, Trade.TypeOfTrade.MIS, "EOD Exit")
                                                        stockList(stockName).EligibleToTakeTrade = False
                                                    End If

                                                    'Stock MTM Check
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    If ExitOnStockFixedTargetStoploss Then
                                                        stockPL = StockPLAfterBrokerage(tradeCheckingDate, runningTick.TradingSymbol)
                                                        If stockPL >= StockMaxProfitPerDay Then
                                                            ExitStockTradesByForce(runningTick, Trade.TypeOfTrade.MIS, "Max Stock Profit reached for the day")
                                                            stockList(stockName).EligibleToTakeTrade = False
                                                        ElseIf stockPL <= Math.Abs(StockMaxLossPerDay) * -1 Then
                                                            ExitStockTradesByForce(runningTick, Trade.TypeOfTrade.MIS, "Max Stock Loss reached for the day")
                                                            stockList(stockName).EligibleToTakeTrade = False
                                                        End If
                                                    End If

                                                    'OverAll MTM Check
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim totalPLOftheDay As Decimal = Decimal.MinValue
                                                    If ExitOnOverAllFixedTargetStoploss Then
                                                        totalPLOftheDay = TotalPLAfterBrokerage(tradeCheckingDate)
                                                        If totalPLOftheDay >= OverAllProfitPerDay Then
                                                            ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Max Profit reached for the day")
                                                            stockList(stockName).EligibleToTakeTrade = False
                                                        ElseIf totalPLOftheDay <= Math.Abs(OverAllLossPerDay) * -1 Then
                                                            ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Max Loss reached for the day")
                                                            stockList(stockName).EligibleToTakeTrade = False
                                                        ElseIf Me.TypeOfMTMTrailing = MTMTrailingType.FixedSlabTrailing AndAlso totalPLOftheDay <= OverAllLossPerDay Then
                                                            ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Trailing MTM reached for the day")
                                                            stockList(stockName).EligibleToTakeTrade = False
                                                        ElseIf Me.TypeOfMTMTrailing = MTMTrailingType.LogSlabTrailing AndAlso totalPLOftheDay <= OverAllLossPerDay Then
                                                            ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Log MTM reached for the day")
                                                            stockList(stockName).EligibleToTakeTrade = False
                                                        ElseIf Me.TypeOfMTMTrailing = MTMTrailingType.RealtimeTrailing AndAlso totalPLOftheDay <= OverAllLossPerDay Then
                                                            ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Realtime Trailing MTM reached for the day")
                                                            stockList(stockName).EligibleToTakeTrade = False
                                                        End If
                                                    End If

                                                    'Exit Trade From Rule
                                                    Dim exitOrderRuleSuccessful As Boolean = False
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim potentialRuleCancelTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Open)
                                                    If potentialRuleCancelTrades IsNot Nothing AndAlso potentialRuleCancelTrades.Count > 0 Then
                                                        If Me.TickBasedStrategy OrElse Not stockList(stockName).CancelOrderDoneForTheMinute Then
                                                            For Each runningCancelTrade In potentialRuleCancelTrades
                                                                _canceller.Token.ThrowIfCancellationRequested()
                                                                Dim exitOrderDetails As Tuple(Of Boolean, String) = Await stockStrategyRule.IsTriggerReceivedForExitOrderAsync(runningTick, runningCancelTrade).ConfigureAwait(False)
                                                                If exitOrderDetails IsNot Nothing AndAlso exitOrderDetails.Item1 Then
                                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                                    ExitTradeByForce(runningCancelTrade, runningTick, exitOrderDetails.Item2)
                                                                    exitOrderRuleSuccessful = True
                                                                End If
                                                            Next
                                                            stockList(stockName).CancelOrderDoneForTheMinute = True
                                                        End If
                                                    End If
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim potentialRuleExitTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
                                                    If potentialRuleExitTrades IsNot Nothing AndAlso potentialRuleExitTrades.Count > 0 Then
                                                        If Me.TickBasedStrategy OrElse Not stockList(stockName).ExitOrderDoneForTheMinute Then
                                                            For Each runningExitTrade In potentialRuleExitTrades
                                                                _canceller.Token.ThrowIfCancellationRequested()
                                                                Dim exitOrderDetails As Tuple(Of Boolean, String) = Await stockStrategyRule.IsTriggerReceivedForExitOrderAsync(runningTick, runningExitTrade).ConfigureAwait(False)
                                                                If exitOrderDetails IsNot Nothing AndAlso exitOrderDetails.Item1 Then
                                                                    ExitTradeByForce(runningExitTrade, runningTick, exitOrderDetails.Item2)
                                                                    exitOrderRuleSuccessful = True
                                                                End If
                                                            Next
                                                            stockList(stockName).ExitOrderDoneForTheMinute = True
                                                        End If
                                                    End If

                                                    'Place Order
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim placeOrderDetails As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
                                                    If stockList(stockName).EligibleToTakeTrade Then
                                                        stockPL = Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol)
                                                        totalPLOftheDay = Me.TotalPLAfterBrokerage(runningTick.PayloadDate)
                                                        Dim nmbrOfTrd As Integer = Me.StockNumberOfTrades(runningTick.PayloadDate, runningTick.TradingSymbol)
                                                        If Me.TickBasedStrategy OrElse Not stockList(stockName).PlaceOrderDoneForTheMinute OrElse exitOrderRuleSuccessful Then
                                                            If nmbrOfTrd < Me.NumberOfTradesPerStockPerDay AndAlso
                                                                totalPLOftheDay < Me.OverAllProfitPerDay AndAlso
                                                                totalPLOftheDay > Me.OverAllLossPerDay AndAlso
                                                                stockPL < Me.StockMaxProfitPerDay AndAlso
                                                                stockPL > Math.Abs(Me.StockMaxLossPerDay) * -1 AndAlso
                                                                stockPL < stockStrategyRule.MaxProfitOfThisStock AndAlso
                                                                stockPL > Math.Abs(stockStrategyRule.MaxLossOfThisStock) * -1 Then
                                                                placeOrderDetails = Await stockStrategyRule.IsTriggerReceivedForPlaceOrderAsync(runningTick).ConfigureAwait(False)
                                                                stockList(stockName).PlaceOrderTrigger = placeOrderDetails
                                                                stockList(stockName).PlaceOrderDoneForTheMinute = True
                                                            End If
                                                        Else
                                                            If nmbrOfTrd < Me.NumberOfTradesPerStockPerDay AndAlso
                                                                totalPLOftheDay < Me.OverAllProfitPerDay AndAlso
                                                                totalPLOftheDay > Me.OverAllLossPerDay AndAlso
                                                                stockPL < Me.StockMaxProfitPerDay AndAlso
                                                                stockPL > Math.Abs(Me.StockMaxLossPerDay) * -1 AndAlso
                                                                stockPL < stockStrategyRule.MaxProfitOfThisStock AndAlso
                                                                stockPL > Math.Abs(stockStrategyRule.MaxLossOfThisStock) * -1 Then
                                                                placeOrderDetails = stockList(stockName).PlaceOrderTrigger
                                                            End If
                                                        End If
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
                                                                Select Case runningOrder.OrderType
                                                                    Case Trade.TypeOfOrder.SL
                                                                        If runningOrder.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                                            If runningTick.High >= runningOrder.EntryPrice Then
                                                                                Continue For
                                                                            End If
                                                                        ElseIf runningOrder.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                                            If runningTick.Low <= runningOrder.EntryPrice Then
                                                                                Continue For
                                                                            End If
                                                                        End If
                                                                    Case Trade.TypeOfOrder.Market
                                                                        runningOrder.EntryPrice = runningTick.Open
                                                                    Case Else
                                                                        Throw New NotImplementedException
                                                                End Select
                                                                Dim runningTrade As Trade = New Trade(originatingStrategy:=Me,
                                                                                                      tradingSymbol:=runningTick.TradingSymbol,
                                                                                                      stockType:=Me.StockType,
                                                                                                      orderType:=runningOrder.OrderType,
                                                                                                      tradingDate:=runningTick.PayloadDate,
                                                                                                      entryDirection:=runningOrder.EntryDirection,
                                                                                                      entryPrice:=runningOrder.EntryPrice,
                                                                                                      entryBuffer:=runningOrder.Buffer,
                                                                                                      squareOffType:=Trade.TypeOfTrade.MIS,
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

                                                    'Exit Trade
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim exitOrderSuccessful As Boolean = False
                                                    Dim potentialExitTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
                                                    If potentialExitTrades IsNot Nothing AndAlso potentialExitTrades.Count > 0 Then
                                                        For Each runningPotentialExitTrade In potentialExitTrades
                                                            _canceller.Token.ThrowIfCancellationRequested()
                                                            exitOrderSuccessful = ExitTradeIfPossible(runningPotentialExitTrade, runningTick)
                                                        Next
                                                    End If

                                                    'Place Order
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim placeOrderTrigger As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
                                                    If exitOrderSuccessful Then
                                                        totalPLOftheDay = Me.TotalPLAfterBrokerage(runningTick.PayloadDate)
                                                        stockPL = Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol)
                                                        Dim nmbrOfTrd As Integer = Me.StockNumberOfTrades(runningTick.PayloadDate, runningTick.TradingSymbol)
                                                        If nmbrOfTrd < Me.NumberOfTradesPerStockPerDay AndAlso
                                                            totalPLOftheDay < Me.OverAllProfitPerDay AndAlso
                                                            totalPLOftheDay > Me.OverAllLossPerDay AndAlso
                                                            stockPL < Me.StockMaxProfitPerDay AndAlso
                                                            stockPL > Math.Abs(Me.StockMaxLossPerDay) * -1 AndAlso
                                                            stockPL < stockStrategyRule.MaxProfitOfThisStock AndAlso
                                                            stockPL > Math.Abs(stockStrategyRule.MaxLossOfThisStock) * -1 Then
                                                            placeOrderTrigger = Await stockStrategyRule.IsTriggerReceivedForPlaceOrderAsync(runningTick).ConfigureAwait(False)
                                                            stockList(stockName).PlaceOrderTrigger = placeOrderTrigger
                                                        End If
                                                    End If
                                                    If placeOrderTrigger IsNot Nothing AndAlso placeOrderTrigger.Item1 Then
                                                        Dim placeOrders As List(Of PlaceOrderParameters) = placeOrderTrigger.Item2
                                                        If placeOrders IsNot Nothing AndAlso placeOrders.Count > 0 Then
                                                            Dim tradeTag As String = System.Guid.NewGuid.ToString()
                                                            For Each runningOrder In placeOrders
                                                                _canceller.Token.ThrowIfCancellationRequested()
                                                                If runningOrder.Used Then
                                                                    Continue For
                                                                End If
                                                                Select Case runningOrder.OrderType
                                                                    Case Trade.TypeOfOrder.SL
                                                                        If runningOrder.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                                            If runningTick.High >= runningOrder.EntryPrice Then
                                                                                Continue For
                                                                            End If
                                                                        ElseIf runningOrder.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                                            If runningTick.Low <= runningOrder.EntryPrice Then
                                                                                Continue For
                                                                            End If
                                                                        End If
                                                                    Case Trade.TypeOfOrder.Market
                                                                        runningOrder.EntryPrice = runningTick.Open
                                                                    Case Else
                                                                        Throw New NotImplementedException
                                                                End Select
                                                                Dim runningTrade As Trade = New Trade(originatingStrategy:=Me,
                                                                                                      tradingSymbol:=runningTick.TradingSymbol,
                                                                                                      stockType:=Me.StockType,
                                                                                                      orderType:=runningOrder.OrderType,
                                                                                                      tradingDate:=runningTick.PayloadDate,
                                                                                                      entryDirection:=runningOrder.EntryDirection,
                                                                                                      entryPrice:=runningOrder.EntryPrice,
                                                                                                      entryBuffer:=runningOrder.Buffer,
                                                                                                      squareOffType:=Trade.TypeOfTrade.MIS,
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
                                                    Dim potentialEntryTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Open)
                                                    If potentialEntryTrades IsNot Nothing AndAlso potentialEntryTrades.Count > 0 Then
                                                        For Each runningPotentialEntryTrade In potentialEntryTrades
                                                            _canceller.Token.ThrowIfCancellationRequested()
                                                            EnterTradeIfPossible(runningPotentialEntryTrade, runningTick)
                                                        Next
                                                    End If

                                                    'Modify Stoploss Trade
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim potentialModifySLTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
                                                    If potentialModifySLTrades IsNot Nothing AndAlso potentialModifySLTrades.Count > 0 Then
                                                        'If Me.TickBasedStrategy OrElse Not stockList(stockName).ModifyStoplossOrderDoneForTheMinute OrElse Me.TrailingStoploss Then
                                                        For Each runningModifyTrade In potentialModifySLTrades
                                                            _canceller.Token.ThrowIfCancellationRequested()
                                                            Dim modifyOrderDetails As Tuple(Of Boolean, Decimal, String) = Await stockStrategyRule.IsTriggerReceivedForModifyStoplossOrderAsync(runningTick, runningModifyTrade).ConfigureAwait(False)
                                                            If modifyOrderDetails IsNot Nothing AndAlso modifyOrderDetails.Item1 Then
                                                                _canceller.Token.ThrowIfCancellationRequested()
                                                                runningModifyTrade.UpdateTrade(PotentialStopLoss:=modifyOrderDetails.Item2, SLRemark:=modifyOrderDetails.Item3)
                                                            End If
                                                        Next
                                                        stockList(stockName).ModifyStoplossOrderDoneForTheMinute = True
                                                        'End If
                                                    End If

                                                    'Modify Target Trade
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim potentialModifyTargetTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
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
                                                Next
                                            End If
                                        End If
                                    Next
                                    startSecond = startSecond.Add(TimeSpan.FromSeconds(1))
                                End While   'Second Loop
                                startMinute = startMinute.Add(TimeSpan.FromMinutes(Me.SignalTimeFrame))
                            End While   'Minute Loop
                            ExitAllTradeByForce(tradeCheckingDate, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Special Force Close")
                        End If
                    End If
                    SetOverallDrawUpDrawDownForTheDay(tradeCheckingDate)
                    totalPL += TotalPLAfterBrokerage(tradeCheckingDate)
                    tradeCheckingDate = tradeCheckingDate.AddDays(1)

                    'Serialization
                    If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                        OnHeartbeat("Serializing Trades collection")
                        SerializeFromCollectionUsingFileStream(Of Dictionary(Of Date, Dictionary(Of String, List(Of Trade))))(tradesFileName, TradesTaken)
                    End If
                End While   'Date Loop

                If CapitalMovement IsNot Nothing Then Utilities.Strings.SerializeFromCollection(Of Dictionary(Of Date, List(Of Capital)))(capitalFileName, CapitalMovement)

                PrintArrayToExcel(filename, tradesFileName, capitalFileName)
            End If
        End Function


#Region "Stock Selection"
        Private Async Function GetStockDataAsync(tradingDate As Date) As Task(Of Dictionary(Of String, StockDetails))
            Dim ret As Dictionary(Of String, StockDetails) = Nothing
            Await Task.Delay(0).ConfigureAwait(False)
            If Me.StockFileName IsNot Nothing Then
                Dim dt As DataTable = Nothing
                Using csvHelper As New Utilities.DAL.CSVHelper(Me.StockFileName, ",", _canceller)
                    dt = csvHelper.GetDataTableFromCSV(1)
                End Using
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Select Case Me.RuleNumber
                        Case 2
                            Dim nifty50Payload As Dictionary(Of Date, Payload) = Cmn.GetRawPayload(Common.DataBaseTable.Intraday_Cash, "NIFTY 50", tradingDate, tradingDate)
                            If nifty50Payload IsNot Nothing AndAlso nifty50Payload.Count > 0 Then
                                Dim nifty50XminPayload As Dictionary(Of Date, Payload) = Common.ConvertPayloadsToXMinutes(nifty50Payload, Me.SignalTimeFrame, New Date(tradingDate.Year, tradingDate.Month, tradingDate.Day, 9, 15, 0))
                                If nifty50XminPayload IsNot Nothing AndAlso nifty50XminPayload.Count > 0 Then
                                    Dim firstCandle As Payload = nifty50XminPayload.FirstOrDefault.Value
                                    If firstCandle.CandleColor = Color.Green OrElse firstCandle.CandleColor = Color.Red Then
                                        Dim counter As Integer = 0
                                        For i = 0 To dt.Rows.Count - 1
                                            Dim rowDate As Date = dt.Rows(i).Item("Date")
                                            If rowDate.Date = tradingDate.Date Then
                                                Dim tradingSymbol As String = dt.Rows(i).Item("Trading Symbol")
                                                Dim instrumentName As String = Nothing
                                                If tradingSymbol.Contains("FUT") Then
                                                    instrumentName = tradingSymbol.Remove(tradingSymbol.Count - 8)
                                                Else
                                                    instrumentName = tradingSymbol
                                                End If
                                                Dim changePer As Decimal = dt.Rows(i).Item("Change %")

                                                Dim lotsize As Integer = dt.Rows(i).Item("Lot Size")
                                                Dim slab As Decimal = dt.Rows(i).Item("Slab")
                                                Dim detailsOfStock As StockDetails = New StockDetails With
                                                {.StockName = instrumentName,
                                                 .TradingSymbol = tradingSymbol,
                                                 .LotSize = lotsize,
                                                 .Slab = slab,
                                                 .EligibleToTakeTrade = True,
                                                 .Supporting1 = changePer,
                                                 .Supporting2 = If(firstCandle.CandleColor = Color.Green, 1, -1)}

                                                If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                                                ret.Add(instrumentName, detailsOfStock)

                                                counter += 1
                                                If counter = Me.NumberOfTradeableStockPerDay Then Exit For
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        Case 4
                            Dim slabList As List(Of Decimal) = New List(Of Decimal) From {0.5, 1, 2.5, 5, 10, 15}
                            Dim counter As Integer = 0
                            For i = 0 To dt.Rows.Count - 1
                                Dim rowDate As Date = dt.Rows(i).Item("Date")
                                If rowDate.Date = tradingDate.Date Then
                                    Dim tradingSymbol As String = dt.Rows(i).Item("Trading Symbol")
                                    Dim instrumentName As String = Nothing
                                    If tradingSymbol.Contains("FUT") Then
                                        instrumentName = tradingSymbol.Remove(tradingSymbol.Count - 8)
                                    Else
                                        instrumentName = tradingSymbol
                                    End If
                                    Dim lotsize As Integer = dt.Rows(i).Item("Lot Size")
                                    Dim slab As Decimal = dt.Rows(i).Item("Slab")
                                    Dim previousDayClose As Decimal = dt.Rows(i).Item("Previous Day Close")
                                    Dim previousSlab As List(Of Decimal) = slabList.FindAll(Function(x)
                                                                                                Return x < slab
                                                                                            End Function)
                                    If previousSlab IsNot Nothing AndAlso previousSlab.Count > 0 Then
                                        Dim projectedSlab As Decimal = previousSlab.LastOrDefault
                                        Dim buffer As Decimal = CalculateBuffer(previousDayClose, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                                        Dim slPoint As Decimal = projectedSlab + 2 * buffer
                                        Dim pl As Decimal = CalculatePL(instrumentName, previousDayClose, previousDayClose - slPoint, lotsize, lotsize, Me.StockType)
                                        If Math.Abs(pl) >= 600 AndAlso Math.Abs(pl) <= 1200 Then
                                            slab = projectedSlab
                                        Else
                                            slab = Decimal.MinValue
                                        End If
                                    Else
                                        slab = Decimal.MinValue
                                    End If

                                    If slab <> Decimal.MinValue Then
                                        Dim detailsOfStock As StockDetails = New StockDetails With
                                                {.StockName = instrumentName,
                                                 .TradingSymbol = tradingSymbol,
                                                 .LotSize = lotsize,
                                                 .Slab = slab,
                                                 .EligibleToTakeTrade = True}

                                        If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                                        ret.Add(instrumentName, detailsOfStock)

                                        counter += 1
                                        If counter = Me.NumberOfTradeableStockPerDay Then Exit For
                                    End If
                                End If
                            Next
                        Case 10
                            Dim counter As Integer = 0
                            For i = 0 To dt.Rows.Count - 1
                                Dim rowDate As Date = dt.Rows(i).Item("Date")
                                If rowDate.Date = tradingDate.Date Then
                                    Dim tradingSymbol As String = dt.Rows(i).Item("Trading Symbol")
                                    Dim instrumentName As String = Nothing
                                    If tradingSymbol.Contains("FUT") Then
                                        instrumentName = tradingSymbol.Remove(tradingSymbol.Count - 8)
                                    Else
                                        instrumentName = tradingSymbol
                                    End If
                                    Dim lotsize As Integer = dt.Rows(i).Item("Lot Size")
                                    Dim slab As Decimal = dt.Rows(i).Item("Slab")
                                    'Dim price As Decimal = dt.Rows(i).Item("Previous Day Close")
                                    'Dim turnover As Decimal = (price * lotsize) / Me.MarginMultiplier

                                    'Console.WriteLine(String.Format("{0}, {1}", tradingSymbol, turnover))
                                    'If turnover < 7000 Then
                                    Dim detailsOfStock As StockDetails = New StockDetails With
                                                {.StockName = instrumentName,
                                                 .TradingSymbol = tradingSymbol,
                                                 .LotSize = lotsize,
                                                 .Slab = slab,
                                                 .EligibleToTakeTrade = True}

                                    If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                                    ret.Add(instrumentName, detailsOfStock)

                                    counter += 1
                                    If counter = Me.NumberOfTradeableStockPerDay Then Exit For
                                    'End If
                                End If
                            Next
                        Case 12
                            Dim stockList As Dictionary(Of String, StockDetails) = Nothing
                            For i = 0 To dt.Rows.Count - 1
                                Dim rowDate As Date = dt.Rows(i).Item("Date")
                                If rowDate.Date = tradingDate.Date Then
                                    Dim tradingSymbol As String = dt.Rows(i).Item("Trading Symbol")
                                    Dim instrumentName As String = Nothing
                                    If tradingSymbol.Contains("FUT") Then
                                        instrumentName = tradingSymbol.Remove(tradingSymbol.Count - 8)
                                    Else
                                        instrumentName = tradingSymbol
                                    End If
                                    Dim lotsize As Integer = dt.Rows(i).Item("Lot Size")
                                    Dim slab As Decimal = dt.Rows(i).Item("Slab")
                                    Dim atr As Decimal = dt.Rows(i).Item("ATR %")
                                    Dim range As Decimal = dt.Rows(i).Item("Volume Per Range")
                                    Dim detailsOfStock As StockDetails = New StockDetails With
                                                {.StockName = instrumentName,
                                                 .TradingSymbol = tradingSymbol,
                                                 .LotSize = lotsize,
                                                 .Slab = slab,
                                                 .EligibleToTakeTrade = True,
                                                 .Supporting1 = atr,
                                                 .Supporting2 = range}

                                    If stockList Is Nothing Then stockList = New Dictionary(Of String, StockDetails)
                                    stockList.Add(instrumentName, detailsOfStock)
                                End If
                            Next
                            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                                Dim avgATR As Decimal = stockList.Average(Function(x)
                                                                              Return x.Value.Supporting1
                                                                          End Function)
                                'Dim minimumATR As Decimal = Math.Floor(avgATR)
                                Dim minimumATR As Decimal = 5
                                Dim supportedStocks As IEnumerable(Of KeyValuePair(Of String, StockDetails)) = stockList.Where(Function(x)
                                                                                                                                   Return x.Value.Supporting1 >= minimumATR
                                                                                                                               End Function)
                                If supportedStocks IsNot Nothing AndAlso supportedStocks.Count > 0 Then
                                    Dim counter As Integer = 0
                                    For Each runningStock In supportedStocks.OrderByDescending(Function(x)
                                                                                                   Return x.Value.Supporting2
                                                                                               End Function)
                                        If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                                        ret.Add(runningStock.Key, runningStock.Value)

                                        counter += 1
                                        If counter >= Me.NumberOfTradeableStockPerDay Then Exit For
                                    Next
                                End If
                            End If
                        Case 15
                            Dim stockDetails As List(Of LowerPriceOptionBuyOnlyStrategyRule.OptionInstumentDetails) = Nothing
                            For i = 0 To dt.Rows.Count - 1
                                Dim rowDate As Date = dt.Rows(i).Item("Date")
                                If rowDate.Date = tradingDate.Date Then
                                    Dim tradingSymbol As String = dt.Rows(i).Item("Trading Symbol")
                                    Dim lotsize As Integer = dt.Rows(i).Item("Lot Size")
                                    Dim instrumentType As String = dt.Rows(i).Item("Puts_Calls")
                                    Dim close As Decimal = dt.Rows(i).Item("Previous Day Close")
                                    Dim volume As Decimal = dt.Rows(i).Item("Previous Day Volume")
                                    Dim oi As Decimal = dt.Rows(i).Item("Previous Day OI")
                                    Dim detailsOfStock As LowerPriceOptionBuyOnlyStrategyRule.OptionInstumentDetails = New LowerPriceOptionBuyOnlyStrategyRule.OptionInstumentDetails With
                                                {.TradingSymbol = tradingSymbol, .LotSize = lotsize, .InstrumentType = instrumentType, .Close = close, .Volume = volume, .OI = oi}

                                    If stockDetails Is Nothing Then stockDetails = New List(Of LowerPriceOptionBuyOnlyStrategyRule.OptionInstumentDetails)
                                    stockDetails.Add(detailsOfStock)
                                End If
                            Next
                            If stockDetails IsNot Nothing AndAlso stockDetails.Count > 0 Then
                                Dim peStocks As List(Of LowerPriceOptionBuyOnlyStrategyRule.OptionInstumentDetails) = stockDetails.FindAll(Function(x)
                                                                                                                                               Return x.InstrumentType = "PE"
                                                                                                                                           End Function)
                                Dim ceStocks As List(Of LowerPriceOptionBuyOnlyStrategyRule.OptionInstumentDetails) = stockDetails.FindAll(Function(x)
                                                                                                                                               Return x.InstrumentType = "CE"
                                                                                                                                           End Function)
                                If peStocks IsNot Nothing AndAlso peStocks.Count > 0 AndAlso ceStocks IsNot Nothing AndAlso ceStocks.Count > 0 Then
                                    Dim avgPEVol As Double = peStocks.OrderByDescending(Function(x)
                                                                                            Return x.Volume
                                                                                        End Function).Take(3).Sum(Function(y)
                                                                                                                      Return y.Volume
                                                                                                                  End Function) / 3
                                    Dim avgCEVol As Double = ceStocks.OrderByDescending(Function(x)
                                                                                            Return x.Volume
                                                                                        End Function).Take(3).Sum(Function(y)
                                                                                                                      Return y.Volume
                                                                                                                  End Function) / 3
                                    Dim avgPEOI As Double = peStocks.OrderByDescending(Function(x)
                                                                                           Return x.OI
                                                                                       End Function).Take(3).Sum(Function(y)
                                                                                                                     Return y.OI
                                                                                                                 End Function) / 3
                                    Dim avgCEOI As Double = ceStocks.OrderByDescending(Function(x)
                                                                                           Return x.OI
                                                                                       End Function).Take(3).Sum(Function(y)
                                                                                                                     Return y.OI
                                                                                                                 End Function) / 3
                                    Dim peStockPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                                    For Each runningStock In peStocks.OrderByDescending(Function(x)
                                                                                            Return x.Volume
                                                                                        End Function)
                                        Dim queryString As String = String.Format("SELECT `Open`,`Low`,`High`,`Close`,`Volume`,`SnapshotDateTime`,`TradingSymbol` 
                                                                                   FROM `intraday_prices_opt_futures` 
                                                                                   WHERE `TradingSymbol` = '{0}' 
                                                                                   AND `SnapshotDate`='{1}'",
                                                                                   runningStock.TradingSymbol, tradingDate.ToString("yyyy-MM-dd"))
                                        Dim tempdt As DataTable = Await Cmn.RunSelectAsync(queryString).ConfigureAwait(False)
                                        Dim intrdayPayload As Dictionary(Of Date, Payload) = Common.ConvertDataTableToPayload(tempdt, 0, 1, 2, 3, 4, 5, 6)
                                        If intrdayPayload IsNot Nothing AndAlso intrdayPayload.Count > 0 Then
                                            If peStockPayload Is Nothing Then peStockPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                            peStockPayload.Add(runningStock.TradingSymbol, intrdayPayload)
                                        End If
                                    Next
                                    Dim peVolStock As LowerPriceOptionBuyOnlyStrategyRule.OptionInstumentDetails = Nothing
                                    For Each runningStock In peStocks.OrderByDescending(Function(x)
                                                                                            Return x.Volume
                                                                                        End Function)
                                        If peStockPayload IsNot Nothing AndAlso peStockPayload.ContainsKey(runningStock.TradingSymbol) Then
                                            Dim intrdayPayload As Dictionary(Of Date, Payload) = peStockPayload(runningStock.TradingSymbol)
                                            If intrdayPayload IsNot Nothing AndAlso intrdayPayload.Count > 0 Then
                                                Dim open As Decimal = intrdayPayload.FirstOrDefault.Value.Open
                                                If open < 10 Then
                                                    Dim vol As Double = (runningStock.Volume / avgPEVol) * 100
                                                    If vol > 80 Then
                                                        If peVolStock Is Nothing OrElse open < peVolStock.LTP Then
                                                            peVolStock = runningStock
                                                            peVolStock.LTP = open
                                                        End If
                                                    Else
                                                        If peVolStock Is Nothing Then
                                                            peVolStock = runningStock
                                                            peVolStock.LTP = open
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                    Dim peOIStock As LowerPriceOptionBuyOnlyStrategyRule.OptionInstumentDetails = Nothing
                                    For Each runningStock In peStocks.OrderByDescending(Function(x)
                                                                                            Return x.Volume
                                                                                        End Function)
                                        If peStockPayload IsNot Nothing AndAlso peStockPayload.ContainsKey(runningStock.TradingSymbol) AndAlso
                                            peVolStock IsNot Nothing AndAlso peVolStock.TradingSymbol <> runningStock.TradingSymbol Then
                                            Dim intrdayPayload As Dictionary(Of Date, Payload) = peStockPayload(runningStock.TradingSymbol)
                                            If intrdayPayload IsNot Nothing AndAlso intrdayPayload.Count > 0 Then
                                                Dim open As Decimal = intrdayPayload.FirstOrDefault.Value.Open
                                                If open < 10 Then
                                                    Dim oi As Double = (runningStock.OI / avgPEOI) * 100
                                                    If oi > 80 Then
                                                        If peOIStock Is Nothing OrElse open < peOIStock.LTP Then
                                                            peOIStock = runningStock
                                                            peOIStock.LTP = open
                                                        End If
                                                    Else
                                                        If peOIStock Is Nothing Then
                                                            peOIStock = runningStock
                                                            peOIStock.LTP = open
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                    Dim ceStockPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                                    For Each runningStock In ceStocks.OrderByDescending(Function(x)
                                                                                            Return x.Volume
                                                                                        End Function)
                                        Dim queryString As String = String.Format("SELECT `Open`,`Low`,`High`,`Close`,`Volume`,`SnapshotDateTime`,`TradingSymbol` 
                                                                                   FROM `intraday_prices_opt_futures` 
                                                                                   WHERE `TradingSymbol` = '{0}' 
                                                                                   AND `SnapshotDate`='{1}'",
                                                                                   runningStock.TradingSymbol, tradingDate.ToString("yyyy-MM-dd"))
                                        Dim tempdt As DataTable = Await Cmn.RunSelectAsync(queryString).ConfigureAwait(False)
                                        Dim intrdayPayload As Dictionary(Of Date, Payload) = Common.ConvertDataTableToPayload(tempdt, 0, 1, 2, 3, 4, 5, 6)
                                        If intrdayPayload IsNot Nothing AndAlso intrdayPayload.Count > 0 Then
                                            If ceStockPayload Is Nothing Then ceStockPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                            ceStockPayload.Add(runningStock.TradingSymbol, intrdayPayload)
                                        End If
                                    Next
                                    Dim ceVolStock As LowerPriceOptionBuyOnlyStrategyRule.OptionInstumentDetails = Nothing
                                    For Each runningStock In ceStocks.OrderByDescending(Function(x)
                                                                                            Return x.Volume
                                                                                        End Function)
                                        If ceStockPayload IsNot Nothing AndAlso ceStockPayload.ContainsKey(runningStock.TradingSymbol) Then
                                            Dim intrdayPayload As Dictionary(Of Date, Payload) = ceStockPayload(runningStock.TradingSymbol)
                                            If intrdayPayload IsNot Nothing AndAlso intrdayPayload.Count > 0 Then
                                                Dim open As Decimal = intrdayPayload.FirstOrDefault.Value.Open
                                                If open < 10 Then
                                                    Dim vol As Double = (runningStock.Volume / avgCEVol) * 100
                                                    If vol > 80 Then
                                                        If ceVolStock Is Nothing OrElse open < ceVolStock.LTP Then
                                                            ceVolStock = runningStock
                                                            ceVolStock.LTP = open
                                                        End If
                                                    Else
                                                        If ceVolStock Is Nothing Then
                                                            ceVolStock = runningStock
                                                            ceVolStock.LTP = open
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                    Dim ceOIStock As LowerPriceOptionBuyOnlyStrategyRule.OptionInstumentDetails = Nothing
                                    For Each runningStock In ceStocks.OrderByDescending(Function(x)
                                                                                            Return x.Volume
                                                                                        End Function)
                                        If ceStockPayload IsNot Nothing AndAlso ceStockPayload.ContainsKey(runningStock.TradingSymbol) AndAlso
                                            ceVolStock IsNot Nothing AndAlso ceVolStock.TradingSymbol <> runningStock.TradingSymbol Then
                                            Dim intrdayPayload As Dictionary(Of Date, Payload) = ceStockPayload(runningStock.TradingSymbol)
                                            If intrdayPayload IsNot Nothing AndAlso intrdayPayload.Count > 0 Then
                                                Dim open As Decimal = intrdayPayload.FirstOrDefault.Value.Open
                                                If open < 10 Then
                                                    Dim oi As Double = (runningStock.OI / avgCEOI) * 100
                                                    If oi > 80 Then
                                                        If ceOIStock Is Nothing OrElse open < ceOIStock.LTP Then
                                                            ceOIStock = runningStock
                                                            ceOIStock.LTP = open
                                                        End If
                                                    Else
                                                        If ceOIStock Is Nothing Then
                                                            ceOIStock = runningStock
                                                            ceOIStock.LTP = open
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next

                                    If peVolStock IsNot Nothing AndAlso peOIStock IsNot Nothing AndAlso ceVolStock IsNot Nothing AndAlso ceOIStock IsNot Nothing Then
                                        Dim avgPrice As Decimal = (peVolStock.LTP + peOIStock.LTP + ceVolStock.LTP + ceOIStock.LTP) / 4
                                        Dim quantity As Integer = Math.Ceiling(CalculateQuantityFromTargetSL("NIFTY", avgPrice, 0, -2500, Trade.TypeOfStock.Futures) / 75)
                                        Dim detailsOfStock1 As StockDetails = New StockDetails With {.StockName = "NIFTY", .TradingSymbol = peVolStock.TradingSymbol, .LotSize = peVolStock.LotSize, .Supporting1 = peVolStock.LTP, .Supporting2 = quantity, .EligibleToTakeTrade = True}
                                        Dim detailsOfStock2 As StockDetails = New StockDetails With {.StockName = "NIFTY", .TradingSymbol = peOIStock.TradingSymbol, .LotSize = peOIStock.LotSize, .Supporting1 = peOIStock.LTP, .Supporting2 = quantity, .EligibleToTakeTrade = True}
                                        Dim detailsOfStock3 As StockDetails = New StockDetails With {.StockName = "NIFTY", .TradingSymbol = ceVolStock.TradingSymbol, .LotSize = ceVolStock.LotSize, .Supporting1 = ceVolStock.LTP, .Supporting2 = quantity, .EligibleToTakeTrade = True}
                                        Dim detailsOfStock4 As StockDetails = New StockDetails With {.StockName = "NIFTY", .TradingSymbol = ceOIStock.TradingSymbol, .LotSize = ceOIStock.LotSize, .Supporting1 = ceOIStock.LTP, .Supporting2 = quantity, .EligibleToTakeTrade = True}
                                        ret = New Dictionary(Of String, StockDetails)
                                        ret.Add(detailsOfStock1.TradingSymbol, detailsOfStock1)
                                        ret.Add(detailsOfStock2.TradingSymbol, detailsOfStock2)
                                        ret.Add(detailsOfStock3.TradingSymbol, detailsOfStock3)
                                        ret.Add(detailsOfStock4.TradingSymbol, detailsOfStock4)

                                        Console.WriteLine(String.Format("{0}, PE VOL", peVolStock.TradingSymbol))
                                        Console.WriteLine(String.Format("{0}, PE OI", peOIStock.TradingSymbol))
                                        Console.WriteLine(String.Format("{0}, CE VOL", ceVolStock.TradingSymbol))
                                        Console.WriteLine(String.Format("{0}, CE OI", ceOIStock.TradingSymbol))
                                    End If
                                End If
                            End If
                        Case 16
                            Dim stockDetails As List(Of LowerPriceOptionOIChangeBuyOnlyStrategyRule.OptionInstumentDetails) = Nothing
                            For i = 0 To dt.Rows.Count - 1
                                Dim rowDate As Date = dt.Rows(i).Item("Date")
                                If rowDate.Date = tradingDate.Date Then
                                    Dim tradingSymbol As String = dt.Rows(i).Item("Trading Symbol")
                                    Dim lotsize As Integer = dt.Rows(i).Item("Lot Size")
                                    Dim instrumentType As String = dt.Rows(i).Item("Instrument Type")
                                    Dim detailsOfStock As LowerPriceOptionOIChangeBuyOnlyStrategyRule.OptionInstumentDetails = New LowerPriceOptionOIChangeBuyOnlyStrategyRule.OptionInstumentDetails With
                                                {.TradingSymbol = tradingSymbol, .LotSize = lotsize, .InstrumentType = instrumentType}

                                    If stockDetails Is Nothing Then stockDetails = New List(Of LowerPriceOptionOIChangeBuyOnlyStrategyRule.OptionInstumentDetails)
                                    stockDetails.Add(detailsOfStock)
                                End If
                            Next
                            If stockDetails IsNot Nothing AndAlso stockDetails.Count > 0 Then
                                OnHeartbeat(String.Format("Getting stock for trading on {0}", tradingDate.ToString("dd-MM-yyyy")))
                                Dim stockPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                                For Each runningStock In stockDetails
                                    Dim queryString As String = String.Format("SELECT `Open`,`Low`,`High`,`Close`,`Volume`,`SnapshotDateTime`,`TradingSymbol` 
                                                                           FROM `intraday_prices_opt_futures` 
                                                                           WHERE `TradingSymbol` = '{0}' 
                                                                           AND `SnapshotDate`>='{1}' AND `SnapshotDate`<='{2}'",
                                                                           runningStock.TradingSymbol, tradingDate.AddDays(-8).ToString("yyyy-MM-dd"), tradingDate.ToString("yyyy-MM-dd"))
                                    Dim tempdt As DataTable = Await Cmn.RunSelectAsync(queryString).ConfigureAwait(False)
                                    Dim intrdayPayload As Dictionary(Of Date, Payload) = Common.ConvertDataTableToPayload(tempdt, 0, 1, 2, 3, 4, 5, 6)
                                    If intrdayPayload IsNot Nothing AndAlso intrdayPayload.Count > 0 Then
                                        If stockPayload Is Nothing Then stockPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                        stockPayload.Add(runningStock.TradingSymbol, intrdayPayload)
                                    End If
                                Next
                                If stockPayload IsNot Nothing AndAlso stockPayload.Count > 0 Then
                                    Dim triggerTime As Dictionary(Of String, Date) = Nothing
                                    For Each runningStock In stockPayload.Keys
                                        Dim trigger As Date = LowerPriceOptionOIChangeBuyOnlyStrategyRule.GetFirstTradeTriggerTime(stockPayload(runningStock), tradingDate)
                                        If triggerTime Is Nothing Then triggerTime = New Dictionary(Of String, Date)
                                        triggerTime.Add(runningStock, trigger)
                                    Next
                                    If triggerTime IsNot Nothing AndAlso triggerTime.Count > 0 Then
                                        For Each runningStock In stockDetails
                                            If triggerTime.ContainsKey(runningStock.TradingSymbol) Then
                                                runningStock.TriggerTime = triggerTime(runningStock.TradingSymbol)
                                            End If
                                        Next
                                        Dim peStock As LowerPriceOptionOIChangeBuyOnlyStrategyRule.OptionInstumentDetails = stockDetails.FindAll(Function(x)
                                                                                                                                                     Return x.InstrumentType = "PE"
                                                                                                                                                 End Function).OrderBy(Function(y)
                                                                                                                                                                           Return y.TriggerTime
                                                                                                                                                                       End Function).FirstOrDefault
                                        Dim ceStock As LowerPriceOptionOIChangeBuyOnlyStrategyRule.OptionInstumentDetails = stockDetails.FindAll(Function(x)
                                                                                                                                                     Return x.InstrumentType = "CE"
                                                                                                                                                 End Function).OrderBy(Function(y)
                                                                                                                                                                           Return y.TriggerTime
                                                                                                                                                                       End Function).FirstOrDefault

                                        Dim detailsOfStock1 As StockDetails = New StockDetails With {.StockName = "NIFTY", .TradingSymbol = peStock.TradingSymbol, .LotSize = peStock.LotSize, .EligibleToTakeTrade = True}
                                        Dim detailsOfStock2 As StockDetails = New StockDetails With {.StockName = "NIFTY", .TradingSymbol = ceStock.TradingSymbol, .LotSize = ceStock.LotSize, .EligibleToTakeTrade = True}
                                        ret = New Dictionary(Of String, StockDetails)
                                        ret.Add(detailsOfStock1.TradingSymbol, detailsOfStock1)
                                        ret.Add(detailsOfStock2.TradingSymbol, detailsOfStock2)
                                    End If
                                End If
                            End If
                        Case 17
                            Dim stockDetails As List(Of LowerPriceOptionBuyOnlyEODStrategyRule.OptionInstumentDetails) = Nothing
                            For i = 0 To dt.Rows.Count - 1
                                Dim rowDate As Date = dt.Rows(i).Item("Date")
                                If rowDate.Date = tradingDate.Date Then
                                    Dim tradingSymbol As String = dt.Rows(i).Item("Trading Symbol")
                                    Dim lotsize As Integer = dt.Rows(i).Item("Lot Size")
                                    Dim instrumentType As String = dt.Rows(i).Item("Instrument Type")
                                    Dim blankCandle As Decimal = dt.Rows(i).Item("Previous Blank Candle%")

                                    Dim detailsOfStock As LowerPriceOptionBuyOnlyEODStrategyRule.OptionInstumentDetails = New LowerPriceOptionBuyOnlyEODStrategyRule.OptionInstumentDetails With
                                                {.TradingSymbol = tradingSymbol, .LotSize = lotsize, .InstrumentType = instrumentType, .BlankCandlePer = blankCandle}

                                    If stockDetails Is Nothing Then stockDetails = New List(Of LowerPriceOptionBuyOnlyEODStrategyRule.OptionInstumentDetails)
                                    stockDetails.Add(detailsOfStock)
                                End If
                            Next
                            If stockDetails IsNot Nothing AndAlso stockDetails.Count > 0 Then
                                Dim peStockList As List(Of LowerPriceOptionBuyOnlyEODStrategyRule.OptionInstumentDetails) = stockDetails.FindAll(Function(x)
                                                                                                                                                     Return x.InstrumentType = "PE" AndAlso x.BlankCandlePer < 20
                                                                                                                                                 End Function)
                                Dim ceStockList As List(Of LowerPriceOptionBuyOnlyEODStrategyRule.OptionInstumentDetails) = stockDetails.FindAll(Function(x)
                                                                                                                                                     Return x.InstrumentType = "CE" AndAlso x.BlankCandlePer < 20
                                                                                                                                                 End Function)
                                Dim tradableStockCount As Integer = Math.Min(peStockList.Count, ceStockList.Count)
                                Dim maxLoss As Decimal = -1000
                                If tradableStockCount <= 3 Then maxLoss = -2000

                                If tradableStockCount > 0 Then
                                    Dim stockCounter As Integer = 0
                                    For Each runningStock In peStockList
                                        Dim detailsOfStock As StockDetails = New StockDetails With
                                                {.StockName = "NIFTY",
                                                 .TradingSymbol = runningStock.TradingSymbol,
                                                 .LotSize = runningStock.LotSize,
                                                 .Supporting1 = maxLoss,
                                                 .EligibleToTakeTrade = True}

                                        If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                                        ret.Add(detailsOfStock.TradingSymbol, detailsOfStock)

                                        stockCounter += 1
                                        If stockCounter >= tradableStockCount Then Exit For
                                    Next
                                    stockCounter = 0
                                    For Each runningStock In ceStockList
                                        Dim detailsOfStock As StockDetails = New StockDetails With
                                                {.StockName = "NIFTY",
                                                 .TradingSymbol = runningStock.TradingSymbol,
                                                 .LotSize = runningStock.LotSize,
                                                 .Supporting1 = maxLoss,
                                                 .EligibleToTakeTrade = True}

                                        If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                                        ret.Add(detailsOfStock.TradingSymbol, detailsOfStock)

                                        stockCounter += 1
                                        If stockCounter >= tradableStockCount Then Exit For
                                    Next
                                End If
                            End If
                        Case 18
                            Dim stockDetails As List(Of EveryMinuteTopGainerLosserHKReversalStrategyRule.InstumentDetails) = Nothing
                            For i = 0 To dt.Rows.Count - 1
                                Dim rowDate As Date = dt.Rows(i).Item("Date")
                                If rowDate.Date = tradingDate.Date Then
                                    Dim tradingSymbol As String = dt.Rows(i).Item("Trading Symbol")
                                    Dim lotsize As Integer = dt.Rows(i).Item("Lot Size")
                                    Dim signalTime As Date = dt.Rows(i).Item("Time")
                                    Dim gainLossPer As Decimal = dt.Rows(i).Item("Gain Loss %")
                                    Dim detailsOfStock As EveryMinuteTopGainerLosserHKReversalStrategyRule.InstumentDetails = New EveryMinuteTopGainerLosserHKReversalStrategyRule.InstumentDetails With
                                                {.TradingSymbol = tradingSymbol, .LotSize = lotsize, .SignalTime = signalTime, .GainLossPer = gainLossPer}

                                    If stockDetails IsNot Nothing AndAlso stockDetails.Count > 0 Then
                                        Dim availableStock = stockDetails.Find(Function(x)
                                                                                   Return x.TradingSymbol = detailsOfStock.TradingSymbol
                                                                               End Function)
                                        If availableStock Is Nothing Then stockDetails.Add(detailsOfStock)
                                    Else
                                        stockDetails = New List(Of EveryMinuteTopGainerLosserHKReversalStrategyRule.InstumentDetails)
                                        stockDetails.Add(detailsOfStock)
                                    End If
                                End If
                            Next
                            If stockDetails IsNot Nothing AndAlso stockDetails.Count > 0 Then
                                OnHeartbeat(String.Format("Getting stock for trading on {0}", tradingDate.ToString("dd-MM-yyyy")))
                                Dim stockPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                                For Each runningStock In stockDetails
                                    Dim intrdayPayload As Dictionary(Of Date, Payload) = Cmn.GetRawPayload(Common.DataBaseTable.Intraday_Cash, runningStock.TradingSymbol, tradingDate.AddDays(-8), tradingDate)
                                    If intrdayPayload IsNot Nothing AndAlso intrdayPayload.Count > 0 Then
                                        If stockPayload Is Nothing Then stockPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                        stockPayload.Add(runningStock.TradingSymbol, intrdayPayload)
                                    End If
                                Next
                                If stockPayload IsNot Nothing AndAlso stockPayload.Count > 0 Then
                                    For Each runningStock In stockDetails
                                        If stockPayload.ContainsKey(runningStock.TradingSymbol) Then
                                            Dim triggerTime As Date = EveryMinuteTopGainerLosserHKReversalStrategyRule.GetFirstTradeTriggerTime(stockPayload(runningStock.TradingSymbol), tradingDate, runningStock.SignalTime, runningStock.GainLossPer)
                                            runningStock.SignalTriggerTime = triggerTime
                                        End If
                                    Next
                                    Dim counter As Integer = 0
                                    For Each runningStock In stockDetails.OrderBy(Function(x)
                                                                                      Return x.SignalTriggerTime
                                                                                  End Function)
                                        Dim detailsOfStock As StockDetails = New StockDetails With
                                                {.StockName = runningStock.TradingSymbol,
                                                 .TradingSymbol = runningStock.TradingSymbol,
                                                 .LotSize = runningStock.LotSize,
                                                 .Supporting1 = runningStock.GainLossPer,
                                                 .SupportingDate = runningStock.SignalTime,
                                                 .EligibleToTakeTrade = True}

                                        If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                                        ret.Add(detailsOfStock.StockName, detailsOfStock)

                                        counter += 1
                                        If counter = Me.NumberOfTradeableStockPerDay Then Exit For
                                    Next
                                End If
                            End If
                        Case Else
                            Dim counter As Integer = 0
                            For i = 0 To dt.Rows.Count - 1
                                Dim rowDate As Date = dt.Rows(i).Item("Date")
                                If rowDate.Date = tradingDate.Date Then
                                    Dim tradingSymbol As String = dt.Rows(i).Item("Trading Symbol")
                                    Dim instrumentName As String = Nothing
                                    If tradingSymbol.Contains("FUT") Then
                                        instrumentName = tradingSymbol.Remove(tradingSymbol.Count - 8)
                                    Else
                                        instrumentName = tradingSymbol
                                    End If
                                    Dim lotsize As Integer = dt.Rows(i).Item("Lot Size")
                                    Dim slab As Decimal = dt.Rows(i).Item("Slab")
                                    Dim detailsOfStock As StockDetails = New StockDetails With
                                                {.StockName = instrumentName,
                                                 .TradingSymbol = tradingSymbol,
                                                 .LotSize = lotsize,
                                                 .Slab = slab,
                                                 .EligibleToTakeTrade = True}

                                    If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                                    ret.Add(instrumentName, detailsOfStock)

                                    counter += 1
                                    If counter = Me.NumberOfTradeableStockPerDay Then Exit For
                                End If
                            Next
                    End Select
                End If
            End If
            Return ret
        End Function

        Public Function GetUniqueInstrumentList(ByVal startDate As Date, ByVal endDate As Date) As List(Of String)
            Dim ret As List(Of String) = Nothing
            If Me.StockFileName IsNot Nothing Then
                Dim dt As DataTable = Nothing
                Using csvHelper As New Utilities.DAL.CSVHelper(Me.StockFileName, ",", _canceller)
                    dt = csvHelper.GetDataTableFromCSV(1)
                End Using
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim tradingDate As Date = startDate
                    While tradingDate <= endDate
                        For i = 0 To dt.Rows.Count - 1
                            Dim rowDate As Date = dt.Rows(i).Item("Date")
                            If rowDate.Date = tradingDate.Date Then
                                Dim tradingSymbol As String = dt.Rows(i).Item("Trading Symbol")
                                'If tradingSymbol.Contains("FUT") Then
                                '    tradingSymbol = tradingSymbol.Remove(tradingSymbol.Count - 8)
                                'End If

                                If ret Is Nothing Then ret = New List(Of String)
                                If Not ret.Contains(tradingSymbol.ToUpper) Then ret.Add(tradingSymbol.ToUpper)
                            End If
                        Next

                        tradingDate = tradingDate.AddDays(1)
                    End While
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
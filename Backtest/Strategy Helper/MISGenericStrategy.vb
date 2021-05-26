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
                       ByVal databaseTable As Common.DataBaseTable,
                       ByVal dataSource As SourceOfData,
                       ByVal initialCapital As Decimal,
                       ByVal usableCapital As Decimal,
                       ByVal minimumEarnedCapitalToWithdraw As Decimal,
                       ByVal amountToBeWithdrawn As Decimal)
            MyBase.New(canceller, exchangeStartTime, exchangeEndTime, tradeStartTime, lastTradeEntryTime, eodExitTime, tickSize, marginMultiplier, timeframe, heikenAshiCandle, stockType, Trade.TypeOfTrade.MIS, databaseTable, dataSource, initialCapital, usableCapital, minimumEarnedCapitalToWithdraw, amountToBeWithdrawn)
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
                    Dim stockList As Dictionary(Of String, StockDetails) = Await GetStockData(tradeCheckingDate).ConfigureAwait(False)

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
                            Dim XDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                            Dim currentDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                            If Me.DataSource = SourceOfData.Database Then
                                XDayOneMinutePayload = Cmn.GetRawPayloadForSpecificTradingSymbol(Me.DatabaseTable, stock, tradeCheckingDate.AddDays(-7), tradeCheckingDate, stockList(stock).OptionStock)
                            ElseIf Me.DataSource = SourceOfData.Live Then
                                Throw New NotImplementedException
                                'XDayOneMinutePayload = Await Cmn.GetHistoricalDataAsync(Me.DatabaseTable, stock, tradeCheckingDate.AddDays(-7), tradeCheckingDate).ConfigureAwait(False)
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
                                    If stocksRuleData Is Nothing Then stocksRuleData = New Dictionary(Of String, StrategyRule)
                                    Dim stockRule As StrategyRule = Nothing

                                    Dim tradingSymbol As String = currentDayOneMinutePayload.LastOrDefault.Value.TradingSymbol
                                    Select Case RuleNumber
                                        Case 0
                                            stockRule = New SaddamOptionStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, Me.RuleEntityData, stockList(stock).Controller, _canceller)
                                        Case Else
                                            Throw New NotImplementedException
                                    End Select

                                    AddHandler stockRule.Heartbeat, AddressOf OnHeartbeat
                                    stockRule.CompletePreProcessing()
                                    stocksRuleData.Add(stock, stockRule)
                                End If
                            End If
                        Next
                        If stocksRuleData IsNot Nothing AndAlso stocksRuleData.Count > 0 Then
                            Dim dependentInstrument As List(Of StrategyRule) = Nothing
                            Dim controller As StrategyRule = Nothing
                            For Each stockRule In stocksRuleData.Values
                                If Not stockRule.Controller Then
                                    If dependentInstrument Is Nothing Then dependentInstrument = New List(Of StrategyRule)
                                    dependentInstrument.Add(stockRule)
                                Else
                                    controller = stockRule
                                End If
                            Next

                            For Each stockRule In stocksRuleData.Values
                                If stockRule.Controller Then
                                    stockRule.DependentInstrument = dependentInstrument
                                    stockRule.CompletePairProcessing()
                                Else
                                    stockRule.ControllerInstrument = controller
                                End If
                            Next
                        End If
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

                                        'Dim pl As Decimal = TotalPLAfterBrokerage(tradeCheckingDate)
                                        'Console.WriteLine(String.Format("{0},{1}", potentialTickSignalTime.AddMinutes(-1).ToString("HH:mm:ss"), pl))
                                    End If
                                    For Each stockName In stockList.Keys
                                        _canceller.Token.ThrowIfCancellationRequested()

                                        If Not stocksRuleData.ContainsKey(stockName) Then
                                            Continue For
                                        End If
                                        Dim stockStrategyRule As StrategyRule = stocksRuleData(stockName)
                                        If Not stockStrategyRule.EligibleToTakeTrade Then
                                            Continue For
                                        End If

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

                                                    ''Update Collection
                                                    'Await stockStrategyRule.UpdateRequiredCollectionsAsync(runningTick).ConfigureAwait(False)

                                                    'Force exit at day end
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    If runningTick.PayloadDate >= eodTime Then
                                                        ExitStockTradesByForce(runningTick, Trade.TypeOfTrade.MIS, "EOD Exit")
                                                        stockList(stockName).EligibleToTakeTrade = False
                                                    End If

                                                    If stockStrategyRule.Controller Then
                                                        Dim totalPLOftheDay As Decimal = TotalPLAfterBrokerage(tradeCheckingDate)

                                                        If Not CType(stockStrategyRule, SaddamOptionStrategyRule).AdjustmentDone Then
                                                            Dim exitDone As Boolean = False
                                                            If totalPLOftheDay >= 3000 Then
                                                                ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Max Profit reached for the day without adjustment")
                                                                exitDone = True
                                                            ElseIf totalPLOftheDay <= -3500 Then
                                                                ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Max Loss reached for the day without adjustment")
                                                                exitDone = True
                                                            End If
                                                            If exitDone Then
                                                                For Each runningStock In stockList
                                                                    runningStock.Value.EligibleToTakeTrade = False
                                                                Next
                                                                Exit For
                                                            End If
                                                        End If

                                                        If totalPLOftheDay >= 5000 Then
                                                            Dim multiplier As Decimal = Math.Floor(totalPLOftheDay / 500)
                                                            Me.ExitOnOverAllFixedTargetStoploss = True
                                                            If (multiplier * 500) - 1000 > Me.OverAllLossPerDay Then
                                                                Me.OverAllLossPerDay = (multiplier * 500) - 1000
                                                            End If
                                                        End If

                                                        'OverAll MTM Check
                                                        _canceller.Token.ThrowIfCancellationRequested()
                                                        If ExitOnOverAllFixedTargetStoploss Then
                                                            Dim exitTrade As Boolean = False
                                                            If totalPLOftheDay >= OverAllProfitPerDay Then
                                                                ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Max Profit reached for the day")
                                                                stockList(stockName).EligibleToTakeTrade = False
                                                                exitTrade = True
                                                            ElseIf totalPLOftheDay <= Math.Abs(OverAllLossPerDay) * -1 Then
                                                                ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Max Loss reached for the day")
                                                                stockList(stockName).EligibleToTakeTrade = False
                                                                exitTrade = True
                                                            ElseIf Me.TypeOfMTMTrailing = MTMTrailingType.FixedSlabTrailing AndAlso totalPLOftheDay <= OverAllLossPerDay Then
                                                                ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Trailing MTM reached for the day")
                                                                stockList(stockName).EligibleToTakeTrade = False
                                                                exitTrade = True
                                                            ElseIf Me.TypeOfMTMTrailing = MTMTrailingType.LogSlabTrailing AndAlso totalPLOftheDay <= OverAllLossPerDay Then
                                                                ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Log MTM reached for the day")
                                                                stockList(stockName).EligibleToTakeTrade = False
                                                                exitTrade = True
                                                            ElseIf Me.TypeOfMTMTrailing = MTMTrailingType.RealtimeTrailing AndAlso totalPLOftheDay <= OverAllLossPerDay Then
                                                                ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.MIS, "Realtime Trailing MTM reached for the day")
                                                                stockList(stockName).EligibleToTakeTrade = False
                                                                exitTrade = True
                                                            End If
                                                            If exitTrade Then
                                                                For Each runningStock In stockList
                                                                    runningStock.Value.EligibleToTakeTrade = False
                                                                Next
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If

                                                    'Exit Trade From Rule
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
                                                                End If
                                                            Next
                                                            stockList(stockName).ExitOrderDoneForTheMinute = True
                                                        End If
                                                    End If
                                                    If stockStrategyRule.Controller Then
                                                        Await stockStrategyRule.IsTriggerReceivedForExitOrderAsync(runningTick, Nothing).ConfigureAwait(False)
                                                    End If

                                                    'Place Order
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim placeOrderDetails As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
                                                    If stockList(stockName).EligibleToTakeTrade Then
                                                        'stockPL = Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol)
                                                        'totalPLOftheDay = Me.TotalPLAfterBrokerage(runningTick.PayloadDate)
                                                        Dim nmbrOfTrd As Integer = Me.StockNumberOfTrades(runningTick.PayloadDate, runningTick.TradingSymbol)
                                                        If Me.TickBasedStrategy OrElse Not stockList(stockName).PlaceOrderDoneForTheMinute Then
                                                            If nmbrOfTrd < Me.NumberOfTradesPerStockPerDay Then
                                                                placeOrderDetails = Await stockStrategyRule.IsTriggerReceivedForPlaceOrderAsync(runningTick).ConfigureAwait(False)
                                                                stockList(stockName).PlaceOrderTrigger = placeOrderDetails
                                                                stockList(stockName).PlaceOrderDoneForTheMinute = True
                                                            End If
                                                        Else
                                                            If nmbrOfTrd < Me.NumberOfTradesPerStockPerDay Then
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
                                                        'totalPLOftheDay = Me.TotalPLAfterBrokerage(runningTick.PayloadDate)
                                                        'stockPL = Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol)
                                                        Dim nmbrOfTrd As Integer = Me.StockNumberOfTrades(runningTick.PayloadDate, runningTick.TradingSymbol)
                                                        If nmbrOfTrd < Me.NumberOfTradesPerStockPerDay Then
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
        Private Async Function GetStockData(tradingDate As Date) As Task(Of Dictionary(Of String, StockDetails))
            Dim ret As Dictionary(Of String, StockDetails) = Nothing
            Await Task.Delay(0).ConfigureAwait(False)
            If Me.StockFileName IsNot Nothing Then
                Dim dt As DataTable = Nothing
                Using csvHelper As New Utilities.DAL.CSVHelper(Me.StockFileName, ",", _canceller)
                    dt = csvHelper.GetDataTableFromCSV(1)
                End Using
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        Dim rowDate As Date = dt.Rows(i).Item("Date")
                        If rowDate.Date = tradingDate.Date Then
                            Dim tradingSymbol As String = dt.Rows(i).Item("Trading Symbol")
                            Dim lotSize As Integer = dt.Rows(i).Item("Lot Size")

                            Dim detailsOfStock As StockDetails = New StockDetails With
                                       {.StockName = tradingSymbol,
                                        .LotSize = lotSize,
                                        .Controller = False,
                                        .OptionStock = True,
                                        .EligibleToTakeTrade = True}

                            If ret Is Nothing Then
                                ret = New Dictionary(Of String, StockDetails)
                                Dim controllerStock As StockDetails = New StockDetails With
                                       {.StockName = "NIFTY BANK",
                                        .LotSize = lotSize,
                                        .Controller = True,
                                        .OptionStock = False,
                                        .EligibleToTakeTrade = True}
                                ret.Add(controllerStock.StockName, controllerStock)
                            End If
                            ret.Add(detailsOfStock.StockName, detailsOfStock)
                        End If
                    Next
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
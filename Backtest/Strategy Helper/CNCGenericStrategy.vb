Imports Algo2TradeBLL
Imports System.IO
Imports System.Threading
Imports Utilities.Strings

Namespace StrategyHelper
    Public Class CNCGenericStrategy
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
                    _canceller.Token.ThrowIfCancellationRequested()
                    Dim stockList As Dictionary(Of String, StockDetails) = GetStockData(tradeCheckingDate)
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
                                XDayOneMinutePayload = Cmn.GetRawPayload(Me.DatabaseTable, stock, tradeCheckingDate.AddDays(-30), tradeCheckingDate)
                            ElseIf Me.DataSource = SourceOfData.Live Then
                                XDayOneMinutePayload = Await Cmn.GetHistoricalDataAsync(Me.DatabaseTable, stock, tradeCheckingDate.AddDays(-20), tradeCheckingDate).ConfigureAwait(False)
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
                                            stockRule = New SwingCNCStrategyRule(XDayOneMinutePayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
                                    End Select

                                    AddHandler stockRule.Heartbeat, AddressOf OnHeartbeat
                                    stockRule.CompletePreProcessing()
                                    stocksRuleData.Add(stock, stockRule)
                                End If
                            End If
                        Next
                        '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                        If currentDayOneMinuteStocksPayload IsNot Nothing AndAlso currentDayOneMinuteStocksPayload.Count > 0 Then
                            OnHeartbeat(String.Format("Checking Trade on {0}", tradeCheckingDate.ToShortDateString))
                            _canceller.Token.ThrowIfCancellationRequested()
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
                                                    SetCurrentLTPForStock(currentMinuteCandlePayload, runningTick, Trade.TypeOfTrade.CNC)

                                                    'Update Collection
                                                    Await stockStrategyRule.UpdateRequiredCollectionsAsync(runningTick).ConfigureAwait(False)

                                                    'Exit Trade From Rule
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim potentialRuleCancelTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Open)
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
                                                    Dim potentialRuleExitTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
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

                                                    'Place Order
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim placeOrderDetails As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
                                                    If Me.TickBasedStrategy OrElse Not stockList(stockName).PlaceOrderDoneForTheMinute Then
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
                                                    Else
                                                        If Me.StockNumberOfTrades(runningTick.PayloadDate, runningTick.TradingSymbol) < Me.NumberOfTradesPerStockPerDay AndAlso
                                                        Me.TotalPLAfterBrokerage(runningTick.PayloadDate) < Me.OverAllProfitPerDay AndAlso
                                                        Me.TotalPLAfterBrokerage(runningTick.PayloadDate) > Math.Abs(Me.OverAllLossPerDay) * -1 AndAlso
                                                        Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol) < Me.StockMaxProfitPerDay AndAlso
                                                        Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol) > Math.Abs(Me.StockMaxLossPerDay) * -1 AndAlso
                                                        Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol) < stockStrategyRule.MaxProfitOfThisStock AndAlso
                                                        Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol) > Math.Abs(stockStrategyRule.MaxLossOfThisStock) * -1 Then
                                                            placeOrderDetails = stockList(stockName).PlaceOrderTrigger
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
                                                                            If runningTick.High > runningOrder.EntryPrice Then
                                                                                Continue For
                                                                            End If
                                                                        ElseIf runningOrder.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                                            If runningTick.Low < runningOrder.EntryPrice Then
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

                                                    'Exit Trade
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim exitOrderSuccessful As Boolean = False
                                                    Dim potentialExitTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
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
                                                        If Me.StockNumberOfTrades(runningTick.PayloadDate, runningTick.TradingSymbol) < Me.NumberOfTradesPerStockPerDay AndAlso
                                                        Me.TotalPLAfterBrokerage(runningTick.PayloadDate) < Me.OverAllProfitPerDay AndAlso
                                                        Me.TotalPLAfterBrokerage(runningTick.PayloadDate) > Math.Abs(Me.OverAllLossPerDay) * -1 AndAlso
                                                        Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol) < Me.StockMaxProfitPerDay AndAlso
                                                        Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol) > Math.Abs(Me.StockMaxLossPerDay) * -1 AndAlso
                                                        Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol) < stockStrategyRule.MaxProfitOfThisStock AndAlso
                                                        Me.StockPLAfterBrokerage(runningTick.PayloadDate, runningTick.TradingSymbol) > Math.Abs(stockStrategyRule.MaxLossOfThisStock) * -1 Then
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
                                                                            If runningTick.High > runningOrder.EntryPrice Then
                                                                                Continue For
                                                                            End If
                                                                        ElseIf runningOrder.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                                            If runningTick.Low < runningOrder.EntryPrice Then
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
                                                    Dim potentialEntryTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Open)
                                                    If potentialEntryTrades IsNot Nothing AndAlso potentialEntryTrades.Count > 0 Then
                                                        For Each runningPotentialEntryTrade In potentialEntryTrades
                                                            _canceller.Token.ThrowIfCancellationRequested()
                                                            EnterTradeIfPossible(runningPotentialEntryTrade, runningTick)
                                                        Next
                                                    End If

                                                    'Modify Stoploss Trade
                                                    _canceller.Token.ThrowIfCancellationRequested()
                                                    Dim potentialModifySLTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
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
                                                    Dim potentialModifyTargetTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
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
                        End If
                        If tradeCheckingDate.Date = endDate.Date Then
                            ExitAllTradeByForce(tradeCheckingDate.AddDays(1).Date, currentDayOneMinuteStocksPayload, Trade.TypeOfTrade.CNC, "Open Trade")
                        End If
                    End If
                    SetOverallDrawUpDrawDownForTheDay(tradeCheckingDate)
                    totalPL += TotalPLAfterBrokerage(tradeCheckingDate)
                    tradeCheckingDate = tradeCheckingDate.AddDays(1)
                End While   'Date Loop

                ''Serialization
                'If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                '    OnHeartbeat("Serializing Trades collection")
                '    SerializeFromCollectionUsingFileStream(Of Dictionary(Of Date, Dictionary(Of String, List(Of Trade))))(tradesFileName, TradesTaken, False)
                'End If
                'If CapitalMovement IsNot Nothing Then Utilities.Strings.SerializeFromCollection(Of Dictionary(Of Date, List(Of Capital)))(capitalFileName, CapitalMovement)

                'PrintArrayToExcel(filename, tradesFileName, capitalFileName)
                PrintArrayToExcel(filename)
            End If
        End Function

#Region "Stock Selection"
        Private Function GetStockData(tradingDate As Date) As Dictionary(Of String, StockDetails)
            Dim ret As Dictionary(Of String, StockDetails) = Nothing

            Dim detailsOfStock As StockDetails = New StockDetails With
                                    {.StockName = "NIFTYBEES",
                                    .LotSize = 1,
                                    .EligibleToTakeTrade = True}

            If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
            ret.Add(detailsOfStock.StockName, detailsOfStock)

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
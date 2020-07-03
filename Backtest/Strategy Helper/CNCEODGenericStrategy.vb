Imports Algo2TradeBLL
Imports System.IO
Imports System.Threading
Imports Utilities.Strings

Namespace StrategyHelper
    Public Class CNCEODGenericStrategy
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
                    'Console.WriteLine(String.Format("Capital available:{0},{1},", tradeCheckingDate.ToShortDateString, Me.AvailableCapital))
                    _canceller.Token.ThrowIfCancellationRequested()
                    Dim stockList As Dictionary(Of String, StockDetails) = GetStockData(tradeCheckingDate, startDate.Date)
                    _canceller.Token.ThrowIfCancellationRequested()
                    If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
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
                                XDayPayload = Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_POSITIONAL, stock, tradeCheckingDate.AddYears(-10), tradeCheckingDate)
                            ElseIf Me.DataSource = SourceOfData.Live Then
                                XDayPayload = Await Cmn.GetHistoricalDataAsync(Me.DatabaseTable, stock, tradeCheckingDate.AddYears(-10), tradeCheckingDate).ConfigureAwait(False)
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
                                            stockRule = New PreviousSwingLowStrategyRule(XDayPayload, stockList(stock).LotSize, Me, tradeCheckingDate, tradingSymbol, _canceller, RuleEntityData)
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

                                    'Place Order
                                    _canceller.Token.ThrowIfCancellationRequested()
                                    Dim placeOrderDetails As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Await stockStrategyRule.IsTriggerReceivedForPlaceOrderAsync(runningTick).ConfigureAwait(False)
                                    stockList(stockName).PlaceOrderTrigger = placeOrderDetails
                                    stockList(stockName).PlaceOrderDoneForTheMinute = True
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
                    tradeCheckingDate = tradeCheckingDate.AddDays(1)
                End While   'Date Loop

                'Serialization
                'If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                '    OnHeartbeat("Serializing Trades collection")
                '    SerializeFromCollectionUsingFileStream(Of Dictionary(Of Date, Dictionary(Of String, List(Of Trade))))(tradesFileName, TradesTaken, False)
                'End If
                'If CapitalMovement IsNot Nothing Then Utilities.Strings.SerializeFromCollection(Of Dictionary(Of Date, List(Of Capital)))(capitalFileName, CapitalMovement)

                'PrintArrayToExcel(filename, tradesFileName, capitalFileName)
                PrintArrayToExcel(filename, Nothing, Nothing)
            End If
        End Function

#Region "Stock Selection"
        Private Function GetStockData(ByVal tradingDate As Date, ByVal tradeStartingDate As Date) As Dictionary(Of String, StockDetails)
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
Imports System.Threading
Imports Utilities.DAL
Imports Algo2TradeBLL
Imports System.IO

Namespace StrategyHelper
    Public MustInherit Class Strategy

#Region "Events/Event handlers"
        Public Event DocumentDownloadComplete()
        Public Event DocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
        Public Event Heartbeat(ByVal msg As String)
        Public Event WaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        'The below functions are needed to allow the derived classes to raise the above two events
        Protected Overridable Sub OnDocumentDownloadComplete()
            RaiseEvent DocumentDownloadComplete()
        End Sub
        Protected Overridable Sub OnDocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
            RaiseEvent DocumentRetryStatus(currentTry, totalTries)
        End Sub
        Protected Overridable Sub OnHeartbeat(ByVal msg As String)
            RaiseEvent Heartbeat(msg)
        End Sub
        Protected Overridable Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
            RaiseEvent WaitingFor(elapsedSecs, totalSecs, msg)
        End Sub
#End Region

#Region "Constructor"
        Public Sub New(ByVal canceller As CancellationTokenSource,
                       ByVal exchangeStartTime As TimeSpan,
                       ByVal exchangeEndTime As TimeSpan,
                       ByVal tradeStartTime As TimeSpan,
                       ByVal tickSize As Decimal,
                       ByVal marginMultiplier As Decimal,
                       ByVal timeframe As Integer,
                       ByVal initialCapital As Decimal,
                       ByVal usableCapital As Decimal,
                       ByVal minimumEarnedCapitalToWithdraw As Decimal,
                       ByVal amountToBeWithdrawn As Decimal,
                       ByVal numberOfLogicalActiveTrade As Integer,
                       ByVal stockFilename As String,
                       ByVal ruleNumber As Integer,
                       ByVal ruleSettings As RuleEntities)
            _canceller = canceller
            Me.ExchangeStartTime = exchangeStartTime
            Me.ExchangeEndTime = exchangeEndTime
            Me.TradeStartTime = tradeStartTime
            Me.TickSize = tickSize
            Me.MarginMultiplier = marginMultiplier
            Me.SignalTimeFrame = timeframe
            Me.InitialCapital = initialCapital
            Me.UsableCapital = usableCapital
            Me.MinimumEarnedCapitalToWithdraw = minimumEarnedCapitalToWithdraw
            Me.AmountToBeWithdrawn = amountToBeWithdrawn
            Me.NumberOfLogicalActiveTrade = numberOfLogicalActiveTrade
            Me.StockFileName = stockFilename
            Me.RuleNumber = ruleNumber
            Me.RuleSettings = ruleSettings

            Cmn = New Common(_canceller)
            'AddHandler Cmn.Heartbeat, AddressOf OnHeartbeat
            'AddHandler Cmn.WaitingFor, AddressOf OnWaitingFor
            'AddHandler Cmn.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
            'AddHandler Cmn.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
        End Sub
#End Region

#Region "Constants"
        Const MARGIN_MULTIPLIER As Integer = 1
        Const DEFAULT_TICK_SIZE As Decimal = 0.05
        Const TRADE_START_TIME As String = "09:15:00"
        Const EXCHANGE_START_TIME As String = "09:15:00"
        Const EXCHANGE_END_TIME As String = "15:29:59"
#End Region

#Region "Variables"
        Protected ReadOnly _canceller As CancellationTokenSource

        Private _TradesTaken As Dictionary(Of String, List(Of Trade))
        Private _CapitalMovement As List(Of Capital) = Nothing
        Private _DayWiseActiveTradeCount As Dictionary(Of Date, Integer) = Nothing
        Private _AvailableCapital As Decimal = Decimal.MinValue
        Private _LogicalActiveTradeCount As Integer = 0

        Public ReadOnly Cmn As Common = Nothing
        Public ReadOnly ExchangeStartTime As TimeSpan = TimeSpan.Parse(EXCHANGE_START_TIME)
        Public ReadOnly ExchangeEndTime As TimeSpan = TimeSpan.Parse(EXCHANGE_END_TIME)
        Public ReadOnly TradeStartTime As TimeSpan = TimeSpan.Parse(TRADE_START_TIME)
        Public ReadOnly TickSize As Decimal = DEFAULT_TICK_SIZE
        Public ReadOnly MarginMultiplier As Decimal = MARGIN_MULTIPLIER
        Public ReadOnly SignalTimeFrame As Integer = 1
        Public ReadOnly InitialCapital As Decimal = Decimal.MaxValue
        Public ReadOnly UsableCapital As Decimal = Decimal.MaxValue
        Public ReadOnly MinimumEarnedCapitalToWithdraw As Decimal = Decimal.MaxValue
        Public ReadOnly AmountToBeWithdrawn As Decimal = Decimal.MaxValue
        Public ReadOnly NumberOfLogicalActiveTrade As Integer = Integer.MaxValue
        Public ReadOnly StockFileName As String = Nothing
        Public ReadOnly RuleNumber As Integer = Integer.MinValue
        Public ReadOnly RuleSettings As RuleEntities = Nothing
#End Region

#Region "Public Functions"
        Public Function GetCurrentXMinuteCandleTime(ByVal time As Date, ByVal timeframe As Integer) As Date
            Dim ret As Date = Nothing
            If Me.ExchangeStartTime.Minutes Mod timeframe = 0 Then
                ret = New Date(time.Year, time.Month, time.Day, time.Hour, Math.Floor(time.Minute / timeframe) * timeframe, 0)
            Else
                Dim exchangeTime As Date = New Date(time.Year, time.Month, time.Day, Me.ExchangeStartTime.Hours, Me.ExchangeStartTime.Minutes, 0)
                Dim currentTime As Date = New Date(time.Year, time.Month, time.Day, time.Hour, time.Minute, 0)
                Dim timeDifference As Double = currentTime.Subtract(exchangeTime).TotalMinutes
                Dim adjustedTimeDifference As Integer = Math.Floor(timeDifference / timeframe) * timeframe
                ret = exchangeTime.AddMinutes(adjustedTimeDifference)
            End If
            Return ret
        End Function

        Public Function GetAllTradesByChildTag(ByVal tag As String, ByVal stockName As String) As List(Of Trade)
            Dim ret As List(Of Trade) = Nothing
            If stockName IsNot Nothing AndAlso _TradesTaken IsNot Nothing AndAlso _TradesTaken.ContainsKey(stockName.ToUpper.Trim) Then
                For Each runningTrade In _TradesTaken(stockName.ToUpper.Trim)
                    If runningTrade.ChildReference = tag Then
                        If ret Is Nothing Then ret = New List(Of Trade)
                        ret.Add(runningTrade)
                    End If
                Next
            End If
            Return ret
        End Function

        Public Function GetAllTradesByParentTag(ByVal tag As String, ByVal stockName As String) As List(Of Trade)
            Dim ret As List(Of Trade) = Nothing
            If stockName IsNot Nothing AndAlso _TradesTaken IsNot Nothing AndAlso _TradesTaken.ContainsKey(stockName.ToUpper.Trim) Then
                For Each runningTrade In _TradesTaken(stockName.ToUpper.Trim)
                    If runningTrade.ParentReference = tag Then
                        If ret Is Nothing Then ret = New List(Of Trade)
                        ret.Add(runningTrade)
                    End If
                Next
            End If
            Return ret
        End Function

        Public Function GetSpecificTrades(ByVal stockName As String, ByVal tradeStatus As Trade.TradeStatus) As List(Of Trade)
            Dim ret As List(Of Trade) = Nothing
            If stockName IsNot Nothing AndAlso _TradesTaken IsNot Nothing AndAlso _TradesTaken.ContainsKey(stockName.ToUpper.Trim) Then
                For Each runningTrade In _TradesTaken(stockName.ToUpper.Trim)
                    If runningTrade.TradeCurrentStatus = tradeStatus Then
                        If ret Is Nothing Then ret = New List(Of Trade)
                        ret.Add(runningTrade)
                    End If
                Next
            End If
            Return ret
        End Function

        Public Function GetLastEntryTradeOfTheStock(ByVal stockName As String) As Trade
            Dim ret As Trade = Nothing
            If stockName IsNot Nothing AndAlso _TradesTaken IsNot Nothing AndAlso _TradesTaken.ContainsKey(stockName.ToUpper.Trim) Then
                ret = _TradesTaken(stockName.ToUpper.Trim).OrderBy(Function(x)
                                                                       Return x.EntryTime
                                                                   End Function).LastOrDefault
            End If
            Return ret
        End Function

        Public Function GetLastCompleteTradeOfTheStock(ByVal stockName As String) As Trade
            Dim ret As Trade = Nothing
            Dim completeTrades As List(Of Trade) = GetSpecificTrades(stockName, Trade.TradeStatus.Complete)
            If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then
                ret = completeTrades.OrderBy(Function(x)
                                                 Return x.EntryTime
                                             End Function).LastOrDefault
            End If
            Return ret
        End Function

        Public Function IsLogicalTradeActiveOfTheStock(ByVal stockName As String) As Boolean
            Dim ret As Boolean = False
            Dim lastEntryTrade As Trade = GetLastEntryTradeOfTheStock(stockName)
            If lastEntryTrade IsNot Nothing AndAlso lastEntryTrade.TradeCurrentStatus = Trade.TradeStatus.Inprogress Then
                ret = True
            Else
                Dim lastCompleteTrade As Trade = GetLastCompleteTradeOfTheStock(stockName)
                If lastCompleteTrade IsNot Nothing AndAlso lastCompleteTrade.ExitType <> Trade.TypeOfExit.Target Then
                    ret = True
                End If
            End If
            Return ret
        End Function

        Public Sub PlaceOrder(ByVal currentTrade As Trade, ByVal currentTick As Payload, ByVal currentTime As Date)
            If currentTrade IsNot Nothing Then
                If Me._AvailableCapital >= currentTrade.CapitalRequiredWithMargin AndAlso
                    _LogicalActiveTradeCount < Me.NumberOfLogicalActiveTrade Then
                    If currentTrade.EntryType = Trade.TypeOfEntry.Fresh Then
                        _LogicalActiveTradeCount += 1
                    End If
                    currentTrade.UpdateTrade(entryTime:=currentTime, tradeCurrentStatus:=Trade.TradeStatus.Open)
                    Dim stockName As String = currentTrade.SpotTradingSymbol.Trim.ToUpper
                    If _TradesTaken Is Nothing Then _TradesTaken = New Dictionary(Of String, List(Of Trade))
                    If _TradesTaken.ContainsKey(stockName) Then
                        _TradesTaken(stockName).Add(currentTrade)
                    Else
                        _TradesTaken.Add(stockName, New List(Of Trade) From {currentTrade})
                    End If

                    InsertCapitalRequired(currentTrade.EntryTime, currentTrade.CapitalRequiredWithMargin, 0, "Place Order")

                    Dim currentMinute As Date = GetCurrentXMinuteCandleTime(currentTime, 1)
                    If currentTick.PayloadDate >= currentMinute AndAlso currentTick.Volume > 0 Then
                        currentTrade.UpdateTrade(entryPrice:=currentTick.Open, entryTime:=currentTime, tradeCurrentStatus:=Trade.TradeStatus.Inprogress)
                    End If
                End If
            End If
        End Sub

        Public Sub ExitOrder(ByVal currentTrade As Trade, ByVal currentTick As Payload, ByVal currentTime As Date, ByVal exitType As Trade.TypeOfExit, ByVal signalCandle As Payload)
            If currentTrade IsNot Nothing AndAlso
                (currentTrade.TradeCurrentStatus = Trade.TradeStatus.Open OrElse
                currentTrade.TradeCurrentStatus = Trade.TradeStatus.Inprogress) Then
                If currentTrade.TradeCurrentStatus = Trade.TradeStatus.Open Then
                    currentTrade.UpdateTrade(exitPrice:=currentTick.Open, exitTime:=currentTime, exitType:=Trade.TypeOfExit.None, exitSignalCandle:=signalCandle, tradeCurrentStatus:=Trade.TradeStatus.Cancel)
                    InsertCapitalRequired(currentTrade.ExitTime, 0, currentTrade.CapitalRequiredWithMargin, "Cancel Order")
                ElseIf currentTrade.TradeCurrentStatus = Trade.TradeStatus.Inprogress Then
                    currentTrade.UpdateTrade(exitPrice:=currentTick.Open, exitTime:=currentTime, exitType:=exitType, exitSignalCandle:=signalCandle, tradeCurrentStatus:=Trade.TradeStatus.Complete)
                    InsertCapitalRequired(currentTrade.ExitTime, 0, currentTrade.CapitalRequiredWithMargin, "Exit Order")
                End If

                If currentTrade.ExitType = Trade.TypeOfExit.Target Then
                    _LogicalActiveTradeCount -= 1
                Else
                    If currentTrade.TradeCurrentStatus = Trade.TradeStatus.Cancel AndAlso
                        currentTrade.EntryType = Trade.TypeOfEntry.Fresh Then
                        _LogicalActiveTradeCount -= 1
                    End If
                End If
            End If
        End Sub

        Public Function CalculatePLAfterBrokerage(ByVal stockName As String, ByVal buyPrice As Decimal, ByVal sellPrice As Decimal, ByVal quantity As Integer, ByVal lotSize As Integer, ByVal typeOfStock As Trade.TypeOfStock) As Decimal
            Dim potentialBrokerage As New Calculator.BrokerageAttributes
            Dim calculator As New Calculator.BrokerageCalculator(_canceller)

            Select Case typeOfStock
                Case Trade.TypeOfStock.Cash
                    calculator.Delivery_Equity(buyPrice, sellPrice, quantity, potentialBrokerage)
                Case Trade.TypeOfStock.Currency
                    calculator.Currency_Futures(buyPrice, sellPrice, quantity / lotSize, potentialBrokerage)
                Case Trade.TypeOfStock.Commodity
                    calculator.Commodity_MCX(stockName, buyPrice, sellPrice, quantity / lotSize, potentialBrokerage)
                Case Trade.TypeOfStock.Futures
                    calculator.FO_Futures(buyPrice, sellPrice, quantity, potentialBrokerage)
                Case Trade.TypeOfStock.Options
                    calculator.FO_Options(buyPrice, sellPrice, quantity, potentialBrokerage)
            End Select

            Return potentialBrokerage.NetProfitLoss
        End Function
#End Region

#Region "Public MustOverride Function"
        Public MustOverride Async Function TestStrategyAsync(startDate As Date, endDate As Date, ByVal filename As String) As Task
#End Region

#Region "Private Functions"
        Private Sub InsertCapitalRequired(ByVal currentDate As Date, ByVal capitalToBeInserted As Decimal, ByVal capitalToBeReleased As Decimal, ByVal remarks As String)
            Dim capitalRequired As Capital = New Capital
            Dim capitalToBeAdded As Decimal = 0

            If _CapitalMovement IsNot Nothing AndAlso _CapitalMovement.Count > 0 Then
                capitalToBeAdded = _CapitalMovement.LastOrDefault.RunningCapital
            Else
                If _CapitalMovement Is Nothing Then _CapitalMovement = New List(Of Capital)
            End If

            Me._AvailableCapital = Me._AvailableCapital - capitalToBeInserted + capitalToBeReleased
            With capitalRequired
                .TradingDate = currentDate.Date
                .CapitalExhaustedDateTime = currentDate
                .RunningCapital = capitalToBeInserted + capitalToBeAdded - capitalToBeReleased
                .CapitalReleased = capitalToBeReleased
                .AvailableCapital = Me._AvailableCapital
                .Remarks = remarks
            End With

            _CapitalMovement.Add(capitalRequired)
        End Sub
#End Region

#Region "Protected Function"
        Protected Sub PopulateDayWiseActiveTradeCount(ByVal currentDay As Date)
            If _DayWiseActiveTradeCount Is Nothing Then _DayWiseActiveTradeCount = New Dictionary(Of Date, Integer)
            _DayWiseActiveTradeCount.Add(currentDay.Date, _LogicalActiveTradeCount)
        End Sub

        Protected Sub SerializeAllCollections(ByVal tradesFilename As String, ByVal capitalFileName As String, ByVal activeTradeFileName As String)
            If _TradesTaken IsNot Nothing AndAlso _TradesTaken.Count > 0 Then
                OnHeartbeat("Serializing Trades collections")
                Utilities.Strings.SerializeFromCollection(Of Dictionary(Of String, List(Of Trade)))(tradesFilename, _TradesTaken)
            End If
            If _CapitalMovement IsNot Nothing AndAlso _CapitalMovement.Count > 0 Then
                OnHeartbeat("Serializing Capital collections")
                Utilities.Strings.SerializeFromCollection(Of List(Of Capital))(capitalFileName, _CapitalMovement)
            End If
            If _DayWiseActiveTradeCount IsNot Nothing AndAlso _DayWiseActiveTradeCount.Count > 0 Then
                OnHeartbeat("Serializing Active Trade collections")
                Utilities.Strings.SerializeFromCollection(Of Dictionary(Of Date, Integer))(activeTradeFileName, _DayWiseActiveTradeCount)
            End If
        End Sub

        Protected Function GetAllActiveStocks() As List(Of StockDetails)
            Dim ret As List(Of StockDetails) = Nothing
            If Me._TradesTaken IsNot Nothing AndAlso Me._TradesTaken.Count > 0 Then
                For Each runningStock In _TradesTaken.Keys
                    _canceller.Token.ThrowIfCancellationRequested()
                    If IsLogicalTradeActiveOfTheStock(runningStock) Then
                        Dim lastEntryTrade As Trade = GetLastEntryTradeOfTheStock(runningStock)
                        Dim stockData As StockDetails = New StockDetails With
                            {.TradingSymbol = runningStock,
                            .LotSize = lastEntryTrade.LotSize}
                        ret.Add(stockData)
                    End If
                Next
            End If
            Return ret
        End Function
#End Region

#Region "Print to excel direct"
        Public Overridable Sub PrintArrayToExcel(ByVal fileName As String, ByVal tradesFilename As String, ByVal capitalFileName As String, ByVal activeTradeFileName As String, ByVal canceller As CancellationTokenSource)
            For retryCounter As Integer = 1 To 20 Step 1
                Try
                    Dim allTradesData As Dictionary(Of String, List(Of Trade)) = Nothing
                    Dim allCapitalData As List(Of Capital) = Nothing
                    Dim allDayWiseActiveTradeCount As Dictionary(Of Date, Integer) = Nothing
                    If tradesFilename IsNot Nothing AndAlso capitalFileName IsNot Nothing AndAlso activeTradeFileName IsNot Nothing AndAlso
                        File.Exists(tradesFilename) AndAlso File.Exists(capitalFileName) AndAlso File.Exists(activeTradeFileName) Then
                        OnHeartbeat("Deserializing Trades collections")
                        allTradesData = Utilities.Strings.DeserializeToCollection(Of Dictionary(Of String, List(Of Trade)))(tradesFilename)
                        canceller.Token.ThrowIfCancellationRequested()
                        OnHeartbeat("Deserializing Capital collections")
                        allCapitalData = Utilities.Strings.DeserializeToCollection(Of List(Of Capital))(capitalFileName)
                        canceller.Token.ThrowIfCancellationRequested()
                        OnHeartbeat("Deserializing Active Trade collections")
                        allDayWiseActiveTradeCount = Utilities.Strings.DeserializeToCollection(Of Dictionary(Of Date, Integer))(activeTradeFileName)

                        If allTradesData IsNot Nothing AndAlso allTradesData.Count > 0 AndAlso
                            allCapitalData IsNot Nothing AndAlso allCapitalData.Count > 0 AndAlso
                            allDayWiseActiveTradeCount IsNot Nothing AndAlso allDayWiseActiveTradeCount.Count > 0 Then
                            OnHeartbeat("Calculating summary")
                            Dim logicalTradeSummary As Dictionary(Of String, Summary) = Nothing
                            If allTradesData IsNot Nothing AndAlso allTradesData.Count > 0 Then
                                For Each runningStock In allTradesData
                                    canceller.Token.ThrowIfCancellationRequested()
                                    Dim uniqueTagList As List(Of String) = New List(Of String)
                                    For Each runningTrade In runningStock.Value
                                        canceller.Token.ThrowIfCancellationRequested()
                                        If Not uniqueTagList.Contains(runningTrade.ParentReference) Then
                                            uniqueTagList.Add(runningTrade.ParentReference)
                                        End If
                                    Next
                                    If uniqueTagList IsNot Nothing AndAlso uniqueTagList.Count > 0 Then
                                        For Each runningTag In uniqueTagList
                                            canceller.Token.ThrowIfCancellationRequested()
                                            Dim tagTrades As List(Of Trade) = runningStock.Value.FindAll(Function(x)
                                                                                                             Return x.ParentReference = runningTag
                                                                                                         End Function)
                                            If tagTrades IsNot Nothing AndAlso tagTrades.Count > 0 Then
                                                If logicalTradeSummary Is Nothing Then logicalTradeSummary = New Dictionary(Of String, Summary)
                                                logicalTradeSummary.Add(runningTag, New Summary With {.AllTrades = tagTrades})
                                            End If
                                        Next
                                    End If
                                Next
                            End If

                            If logicalTradeSummary IsNot Nothing AndAlso logicalTradeSummary.Count > 0 Then
                                OnHeartbeat("Calculating supporting values")
                                Dim totalPL As Decimal = logicalTradeSummary.Values.Sum(Function(x)
                                                                                            Return x.OverallPL
                                                                                        End Function)
                                Dim tradeCount As Integer = logicalTradeSummary.Count
                                Dim maxCapital As Decimal = allCapitalData.Max(Function(x)
                                                                                   Return x.RunningCapital
                                                                               End Function)

                                If tradeCount = 0 Then
                                    fileName = String.Format("PL {0},Cap {1},ROI {2},LgclTrd {3},MaxDays {4},{5}.xlsx",
                                                 Math.Round(totalPL, 0),
                                                 "∞",
                                                 "∞",
                                                 tradeCount,
                                                 "∞",
                                                 fileName)
                                Else
                                    Dim roi As Decimal = (totalPL / maxCapital) * 100
                                    fileName = String.Format("PL {0},Cap {1},ROI {2},LgclTrd {3},{4}.xlsx",
                                                 Math.Round(totalPL, 0),
                                                 Math.Round(maxCapital, 0),
                                                 Math.Round(roi, 0),
                                                 tradeCount,
                                                 fileName)
                                End If

                                Dim filepath As String = Path.Combine(My.Application.Info.DirectoryPath, "BackTest Output", fileName)
                                If File.Exists(filepath) Then File.Delete(filepath)

                                OnHeartbeat("Opening Excel.....")
                                Using excelWriter As New ExcelHelper(filepath, ExcelHelper.ExcelOpenStatus.OpenAfreshForWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, canceller)
                                    excelWriter.CreateNewSheet("Capital")
                                    excelWriter.CreateNewSheet("Data")
                                    excelWriter.SetActiveSheet("Data")

                                    Dim rowCtr As Integer = 0
                                    Dim colCtr As Integer = 0

                                    Dim rowCount As Integer = 0
                                    If allTradesData IsNot Nothing AndAlso allTradesData.Count > 0 Then
                                        rowCount = allTradesData.Sum(Function(x)
                                                                         Dim trades = x.Value.FindAll(Function(z)
                                                                                                          Return z.TradeCurrentStatus <> Trade.TradeStatus.Cancel
                                                                                                      End Function)
                                                                         Return trades.Count
                                                                     End Function)
                                    End If

                                    Dim mainRawData(rowCount, 0) As Object

                                    If rowCtr = 0 Then
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Trading Date"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Trading Symbol"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Capital With Margin"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Signal Direction"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Entry Direction"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Entry Type"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Buy Price"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Sell Price"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Quantity"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Entry Time"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Exit Time"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Exit Type"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "PL Point"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "PL Before Brokerage"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "PL After Brokerage"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Max Draw Up"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Max Draw Down"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Signal Candle Time"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Month"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Child Tag"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Parent Tag"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Iteration Number"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Spot Price"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Spot ATR"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Previous Loss"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Potential Target"
                                        colCtr += 1
                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                        mainRawData(rowCtr, colCtr) = "Spot Instrument"

                                        rowCtr += 1
                                    End If
                                    Dim stockCtr As Integer = 0
                                    For Each stock In allTradesData
                                        canceller.Token.ThrowIfCancellationRequested()
                                        stockCtr += 1
                                        Dim tradeList As List(Of Trade) = stock.Value.FindAll(Function(x)
                                                                                                  Return x.TradeCurrentStatus <> Trade.TradeStatus.Cancel
                                                                                              End Function)
                                        Dim tradeCtr As Integer = 0
                                        If tradeList IsNot Nothing AndAlso tradeList.Count > 0 Then
                                            For Each tradeTaken In tradeList.OrderBy(Function(x)
                                                                                         Return x.EntryTime
                                                                                     End Function)
                                                canceller.Token.ThrowIfCancellationRequested()
                                                tradeCtr += 1
                                                OnHeartbeat(String.Format("Data sheet printing: Stock #{0}/{1} Trade #{2}/{3} #1/3", stockCtr, allTradesData.Count, tradeCtr, tradeList.Count))
                                                colCtr = 0

                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.TradingDate.ToString("dd-MMM-yyyy")
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.TradingSymbol
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.CapitalRequiredWithMargin
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.SignalDirection.ToString
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.EntryDirection.ToString
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.EntryType.ToString
                                                colCtr += 1
                                                If tradeTaken.EntryDirection = Trade.TradeDirection.Buy Then
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.EntryPrice, 4)
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.ExitPrice, 4)
                                                    colCtr += 1
                                                ElseIf tradeTaken.EntryDirection = Trade.TradeDirection.Sell Then
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.ExitPrice, 4)
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.EntryPrice, 4)
                                                    colCtr += 1
                                                End If
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.Quantity
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.EntryTime.ToString("yyyy-MM-dd HH:mm:ss")
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.ExitTime.ToString("yyyy-MM-dd HH:mm:ss")
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.ExitType.ToString
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.PLPoint, 4)
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.PLBeforeBrokerage, 4)
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.PLAfterBrokerage
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.MaxDrawUpPL
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.MaxDrawDownPL
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.EntrySignalCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss")
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = String.Format("{0}-{1}", tradeTaken.TradingDate.ToString("yyyy"), tradeTaken.TradingDate.ToString("MM"))
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.ChildReference
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.ParentReference
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.IterationNumber
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.SpotPrice
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.SpotATR
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.PreviousLoss
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.PotentialTarget
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.SpotTradingSymbol

                                                rowCtr += 1
                                            Next
                                        End If
                                    Next

                                    canceller.Token.ThrowIfCancellationRequested()
                                    Dim range As String = excelWriter.GetNamedRange(1, rowCount, 1, colCtr)
                                    RaiseEvent Heartbeat("Writing from memory to excel...")
                                    excelWriter.WriteArrayToExcel(mainRawData, range)
                                    Erase mainRawData
                                    mainRawData = Nothing

                                    canceller.Token.ThrowIfCancellationRequested()
                                    excelWriter.SetActiveSheet("Capital")
                                    rowCtr = 0
                                    colCtr = 0

                                    rowCount = allCapitalData.Count

                                    Dim capitalRawData(rowCount, 0) As Object

                                    If rowCtr = 0 Then
                                        If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                        capitalRawData(rowCtr, colCtr) = "Trading Date"
                                        colCtr += 1
                                        If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                        capitalRawData(rowCtr, colCtr) = "Capital Exhausted Date Time"
                                        colCtr += 1
                                        If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                        capitalRawData(rowCtr, colCtr) = "Running Capital"
                                        colCtr += 1
                                        If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                        capitalRawData(rowCtr, colCtr) = "Capital Released"
                                        colCtr += 1
                                        If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                        capitalRawData(rowCtr, colCtr) = "Available Capital"
                                        colCtr += 1
                                        If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                        capitalRawData(rowCtr, colCtr) = "Remarks"

                                        rowCtr += 1
                                    End If

                                    For Each runningData In allCapitalData
                                        canceller.Token.ThrowIfCancellationRequested()
                                        OnHeartbeat(String.Format("Capital sheet printing: #{0}/{1} #2/3", rowCount, allCapitalData.Count))
                                        colCtr = 0

                                        If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                        capitalRawData(rowCtr, colCtr) = runningData.TradingDate.ToString("dd-MMM-yyyy")
                                        colCtr += 1
                                        If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                        capitalRawData(rowCtr, colCtr) = runningData.CapitalExhaustedDateTime.ToString("HH:mm:ss")
                                        colCtr += 1
                                        If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(mainRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                        capitalRawData(rowCtr, colCtr) = runningData.RunningCapital
                                        colCtr += 1
                                        If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                        capitalRawData(rowCtr, colCtr) = runningData.CapitalReleased
                                        colCtr += 1
                                        If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                        capitalRawData(rowCtr, colCtr) = runningData.AvailableCapital
                                        colCtr += 1
                                        If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                        capitalRawData(rowCtr, colCtr) = runningData.Remarks

                                        rowCtr += 1
                                    Next

                                    canceller.Token.ThrowIfCancellationRequested()
                                    range = excelWriter.GetNamedRange(1, rowCount, 1, colCtr)
                                    OnHeartbeat("Writing from memory to excel...")
                                    excelWriter.WriteArrayToExcel(capitalRawData, range)
                                    Erase capitalRawData
                                    capitalRawData = Nothing

                                    excelWriter.CreateNewSheet("Active Trade")
                                    excelWriter.SetActiveSheet("Active Trade")

                                    excelWriter.SetData(1, 1, "Date")
                                    excelWriter.SetData(1, 2, "Active Trade Count")

                                    Dim rowNumber As Integer = 1
                                    For Each runningDay In allDayWiseActiveTradeCount
                                        canceller.Token.ThrowIfCancellationRequested()
                                        rowNumber += 1
                                        excelWriter.SetData(rowNumber, 1, runningDay.Key.ToString("dd-MMM-yyyy"))
                                        excelWriter.SetData(rowNumber, 2, runningDay.Value, "######0", ExcelHelper.XLAlign.Right)
                                    Next

                                    If logicalTradeSummary IsNot Nothing AndAlso logicalTradeSummary.Count > 0 Then
                                        excelWriter.CreateNewSheet("Summary")
                                        excelWriter.SetActiveSheet("Summary")

                                        Dim maxTradeCount As Integer = logicalTradeSummary.Max(Function(x)
                                                                                                   Return x.Value.TradeCount
                                                                                               End Function)

                                        Dim colNumber As Integer = 1
                                        excelWriter.SetData(1, colNumber, "Instrument")
                                        colNumber += 1
                                        excelWriter.SetData(1, colNumber, "Start Date")
                                        colNumber += 1
                                        excelWriter.SetData(1, colNumber, "End Date")
                                        colNumber += 1
                                        excelWriter.SetData(1, colNumber, "Number Of Days")
                                        colNumber += 1
                                        excelWriter.SetData(1, colNumber, "Contract Rollover Trade Count")
                                        colNumber += 1
                                        excelWriter.SetData(1, colNumber, "Reverse Signal Trade Count")
                                        colNumber += 1
                                        excelWriter.SetData(1, colNumber, "Overall PL")
                                        colNumber += 1
                                        excelWriter.SetData(1, colNumber, "Max Capital Required")
                                        colNumber += 1
                                        excelWriter.SetData(1, colNumber, "Absolute ROI %")
                                        colNumber += 1
                                        excelWriter.SetData(1, colNumber, "Annual ROI %")
                                        colNumber += 1
                                        excelWriter.SetData(1, colNumber, "Reference")
                                        colNumber += 1
                                        For ctr As Integer = 1 To maxTradeCount
                                            canceller.Token.ThrowIfCancellationRequested()
                                            excelWriter.SetData(1, colNumber, "Contract")
                                            colNumber += 1
                                            excelWriter.SetData(1, colNumber, "Entry Time")
                                            colNumber += 1
                                            excelWriter.SetData(1, colNumber, "Entry Price")
                                            colNumber += 1
                                            excelWriter.SetData(1, colNumber, "Quantity")
                                            colNumber += 1
                                            excelWriter.SetData(1, colNumber, "Exit Time")
                                            colNumber += 1
                                            excelWriter.SetData(1, colNumber, "Exit Price")
                                            colNumber += 1
                                            excelWriter.SetData(1, colNumber, "Exit Reason")
                                            colNumber += 1
                                            excelWriter.SetData(1, colNumber, "PL")
                                            colNumber += 1
                                        Next

                                        rowNumber = 1
                                        For Each runningSummary In logicalTradeSummary
                                            canceller.Token.ThrowIfCancellationRequested()
                                            rowNumber += 1

                                            OnHeartbeat(String.Format("Summary sheet printing: #{0}/{1} #3/3", rowNumber, logicalTradeSummary.Count))

                                            colNumber = 1
                                            excelWriter.SetData(rowNumber, colNumber, runningSummary.Value.Instrument)
                                            colNumber += 1
                                            excelWriter.SetData(rowNumber, colNumber, runningSummary.Value.StartDate.ToString("dd-MMM-yyyy"))
                                            colNumber += 1
                                            excelWriter.SetData(rowNumber, colNumber, runningSummary.Value.EndDate.ToString("dd-MMM-yyyy"))
                                            colNumber += 1
                                            excelWriter.SetData(rowNumber, colNumber, runningSummary.Value.NumberOfDays, "##,##,##0", ExcelHelper.XLAlign.Right)
                                            colNumber += 1
                                            excelWriter.SetData(rowNumber, colNumber, runningSummary.Value.ContractRolloverTradeCount, "##,##,##0", ExcelHelper.XLAlign.Right)
                                            colNumber += 1
                                            excelWriter.SetData(rowNumber, colNumber, runningSummary.Value.ReverseTradeCount, "##,##,##0", ExcelHelper.XLAlign.Right)
                                            colNumber += 1
                                            excelWriter.SetData(rowNumber, colNumber, runningSummary.Value.OverallPL, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                            colNumber += 1
                                            excelWriter.SetData(rowNumber, colNumber, runningSummary.Value.MaxCapital, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                            colNumber += 1
                                            excelWriter.SetData(rowNumber, colNumber, runningSummary.Value.AbsoluteReturnOfInvestment, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                            colNumber += 1
                                            excelWriter.SetData(rowNumber, colNumber, runningSummary.Value.AnnualReturnOfInvestment, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                            colNumber += 1
                                            excelWriter.SetData(rowNumber, colNumber, runningSummary.Key)
                                            colNumber += 1
                                            For Each runningTrade In runningSummary.Value.AllTrades.OrderBy(Function(x)
                                                                                                                Return x.EntryTime
                                                                                                            End Function)
                                                If runningTrade.TradeCurrentStatus <> Trade.TradeStatus.Cancel Then
                                                    excelWriter.SetData(rowNumber, colNumber, runningTrade.TradingSymbol)
                                                    colNumber += 1
                                                    excelWriter.SetData(rowNumber, colNumber, runningTrade.EntryTime.ToString("dd-MMM-yyyy HH:mm:ss"))
                                                    colNumber += 1
                                                    excelWriter.SetData(rowNumber, colNumber, runningTrade.EntryPrice, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                                    colNumber += 1
                                                    excelWriter.SetData(rowNumber, colNumber, runningTrade.Quantity, "######0", ExcelHelper.XLAlign.Right)
                                                    colNumber += 1
                                                    excelWriter.SetData(rowNumber, colNumber, runningTrade.ExitTime.ToString("dd-MMM-yyyy HH:mm:ss"))
                                                    colNumber += 1
                                                    excelWriter.SetData(rowNumber, colNumber, runningTrade.ExitPrice, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                                    colNumber += 1
                                                    excelWriter.SetData(rowNumber, colNumber, runningTrade.ExitType.ToString)
                                                    colNumber += 1
                                                    excelWriter.SetData(rowNumber, colNumber, runningTrade.PLAfterBrokerage, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                                    colNumber += 1
                                                End If
                                            Next
                                        Next
                                    End If

                                    OnHeartbeat("Saving excel...")
                                    excelWriter.SaveExcel()
                                End Using
                            End If

                            File.Delete(tradesFilename)
                            File.Delete(capitalFileName)
                            File.Delete(activeTradeFileName)
                            Exit For
                        Else
                            Exit For
                        End If
                    End If
                Catch ex As System.OutOfMemoryException
                    retryCounter += 1
                    GC.Collect()
                End Try
                Thread.Sleep(30000)
            Next
        End Sub
#End Region

    End Class
End Namespace
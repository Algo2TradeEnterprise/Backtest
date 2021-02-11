Imports Algo2TradeBLL
Imports System.IO
Imports System.Threading
Imports Utilities.Strings

Namespace StrategyHelper
    Public Class CNCGenericStrategy
        Inherits Strategy
        Implements IDisposable



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
            MyBase.New(canceller, exchangeStartTime, exchangeEndTime, tradeStartTime, tickSize, marginMultiplier, timeframe, initialCapital, usableCapital, minimumEarnedCapitalToWithdraw, amountToBeWithdrawn, numberOfLogicalActiveTrade, stockFilename, ruleNumber, ruleSettings)
        End Sub

        Public Overrides Async Function TestStrategyAsync(startDate As Date, endDate As Date, filename As String) As Task
            If filename Is Nothing Then Throw New ApplicationException("Filename Invalid")
            Dim tradesFileName As String = Path.Combine(My.Application.Info.DirectoryPath, String.Format("{0}.Trades.a2t", filename))
            Dim capitalFileName As String = Path.Combine(My.Application.Info.DirectoryPath, String.Format("{0}.Capital.a2t", filename))
            Dim activeTradeFileName As String = Path.Combine(My.Application.Info.DirectoryPath, String.Format("{0}.ActiveTrade.a2t", filename))

            If File.Exists(tradesFileName) AndAlso File.Exists(capitalFileName) AndAlso File.Exists(activeTradeFileName) Then
                PrintArrayToExcel(filename, tradesFileName, capitalFileName, activeTradeFileName, _canceller)
            Else
                If File.Exists(tradesFileName) Then File.Delete(tradesFileName)
                If File.Exists(capitalFileName) Then File.Delete(capitalFileName)
                If File.Exists(activeTradeFileName) Then File.Delete(activeTradeFileName)

                Dim tradeCheckingDate As Date = startDate.Date
                While tradeCheckingDate <= endDate.Date
                    _canceller.Token.ThrowIfCancellationRequested()
                    Dim stockList As List(Of StockDetails) = GetStockData(tradeCheckingDate)
                    _canceller.Token.ThrowIfCancellationRequested()
                    OnHeartbeat("Adding Running stocks")
                    Dim activeStocks As List(Of StockDetails) = GetAllActiveStocks()
                    If activeStocks IsNot Nothing AndAlso activeStocks.Count > 0 Then
                        For Each runningStock In activeStocks
                            If stockList Is Nothing Then stockList = New List(Of StockDetails)
                            Dim availableStock As StockDetails = stockList.Find(Function(x)
                                                                                    Return x.TradingSymbol.ToUpper.Trim = runningStock.TradingSymbol.ToUpper.Trim
                                                                                End Function)
                            If availableStock Is Nothing Then
                                stockList.Add(runningStock)
                            End If
                        Next
                    End If
                    If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                        Dim checkForEntryExit As Boolean = False
                        Dim stocksRuleData As Dictionary(Of String, StrategyRule) = Nothing
                        Dim nextTradingDay As Date = Cmn.GetNexTradingDay(Common.DataBaseTable.EOD_POSITIONAL, tradeCheckingDate)
                        If nextTradingDay = Date.MinValue Then nextTradingDay = Now.Date
                        Dim stkCtr As Integer = 0
                        For Each runningPair In stockList
                            _canceller.Token.ThrowIfCancellationRequested()
                            stkCtr += 1
                            OnHeartbeat(String.Format("Getting Data from DB for {0} on {1} #{2}/{3}", runningPair.TradingSymbol, tradeCheckingDate.ToString("dd-MMM-yyyy"), stkCtr, stockList.Count))
                            Dim XDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                            XDayOneMinutePayload = Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Cash, runningPair.TradingSymbol, tradeCheckingDate.AddMonths(-6), tradeCheckingDate)
                            _canceller.Token.ThrowIfCancellationRequested()
                            Dim XDayEODPayload As Dictionary(Of Date, Payload) = Nothing
                            XDayEODPayload = Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_POSITIONAL, runningPair.TradingSymbol, tradeCheckingDate.AddYears(-1), tradeCheckingDate)
                            _canceller.Token.ThrowIfCancellationRequested()
                            If XDayOneMinutePayload IsNot Nothing AndAlso XDayOneMinutePayload.Count > 0 Then
                                If XDayEODPayload IsNot Nothing AndAlso XDayEODPayload.Count > 100 AndAlso
                                    XDayEODPayload.ContainsKey(tradeCheckingDate.Date) Then
                                    OnHeartbeat(String.Format("Calculating indicators for {0} on {1} #{2}/{3}", runningPair.TradingSymbol, tradeCheckingDate.ToString("dd-MMM-yyyy"), stkCtr, stockList.Count))
                                    If stocksRuleData Is Nothing Then stocksRuleData = New Dictionary(Of String, StrategyRule)
                                    Dim stockRule As StrategyRule = Nothing
                                    Select Case Me.RuleNumber
                                        Case 0
                                            stockRule = New PivotTrendOptionBuyMode3StrategyRule(_canceller, tradeCheckingDate, nextTradingDay, runningPair.TradingSymbol, runningPair.LotSize, Me, XDayOneMinutePayload, XDayEODPayload)
                                        Case Else
                                            Throw New NotImplementedException
                                    End Select

                                    AddHandler stockRule.Heartbeat, AddressOf OnHeartbeat
                                    stockRule.CompletePreProcessing()
                                    stocksRuleData.Add(runningPair.TradingSymbol, stockRule)

                                    checkForEntryExit = True
                                    'Else
                                    '    If Cmn.IsTradingDay(tradeCheckingDate) Then
                                    '        Throw New NotImplementedException(String.Format("Data unavailable for {0}, {1}", runningPair.TradingSymbol, tradeCheckingDate.ToString("dd-MMM-yyyy")))
                                    '    End If
                                End If
                                'Else
                                '    If Cmn.IsTradingDay(tradeCheckingDate) Then
                                '        Throw New NotImplementedException(String.Format("Data unavailable for {0}, {1}", runningPair.TradingSymbol, tradeCheckingDate.ToString("dd-MMM-yyyy")))
                                '    End If
                            End If
                        Next
                        '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                        If checkForEntryExit Then
                            OnHeartbeat(String.Format("Checking Trade on {0}", tradeCheckingDate.ToShortDateString))
                            _canceller.Token.ThrowIfCancellationRequested()
                            Dim startMinute As TimeSpan = Me.ExchangeStartTime
                            Dim endMinute As TimeSpan = Me.ExchangeEndTime
                            While startMinute <= endMinute
                                _canceller.Token.ThrowIfCancellationRequested()
                                OnHeartbeat(String.Format("Checking Trade on {0}. Time:{1}", tradeCheckingDate.ToShortDateString, startMinute.ToString))
                                Dim currentTickSignalTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startMinute.Hours, startMinute.Minutes, startMinute.Seconds)

                                For Each runningStock In stockList
                                    _canceller.Token.ThrowIfCancellationRequested()
                                    Dim stockStrategyRule As StrategyRule = stocksRuleData(runningStock.TradingSymbol)

                                    _canceller.Token.ThrowIfCancellationRequested()
                                    Await stockStrategyRule.UpdateCollectionsIfRequiredAsync(currentTickSignalTime).ConfigureAwait(False)

                                    _canceller.Token.ThrowIfCancellationRequested()
                                    Dim potentialCancelTrades As List(Of Trade) = GetSpecificTrades(runningStock.TradingSymbol, Trade.TradeStatus.Open)
                                    If potentialCancelTrades IsNot Nothing AndAlso potentialCancelTrades.Count > 0 Then
                                        Dim cancelTriggers As List(Of Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload)) = Nothing
                                        cancelTriggers = Await stockStrategyRule.IsTriggerReceivedForExitOrderAsync(currentTickSignalTime, potentialCancelTrades).ConfigureAwait(False)
                                        If cancelTriggers IsNot Nothing AndAlso cancelTriggers.Count > 0 Then
                                            For Each runningTrigger In cancelTriggers
                                                _canceller.Token.ThrowIfCancellationRequested()
                                                ExitOrder(runningTrigger.Item1, runningTrigger.Item2, currentTickSignalTime, runningTrigger.Item3, runningTrigger.Item4)
                                            Next
                                        End If
                                    End If

                                    _canceller.Token.ThrowIfCancellationRequested()
                                    Dim potentialExitTrades As List(Of Trade) = GetSpecificTrades(runningStock.TradingSymbol, Trade.TradeStatus.Inprogress)
                                    If potentialExitTrades IsNot Nothing AndAlso potentialExitTrades.Count > 0 Then
                                        Dim exitTriggers As List(Of Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload)) = Nothing
                                        exitTriggers = Await stockStrategyRule.IsTriggerReceivedForExitOrderAsync(currentTickSignalTime, potentialExitTrades).ConfigureAwait(False)
                                        If exitTriggers IsNot Nothing AndAlso exitTriggers.Count > 0 Then
                                            For Each runningTrigger In exitTriggers
                                                _canceller.Token.ThrowIfCancellationRequested()
                                                ExitOrder(runningTrigger.Item1, runningTrigger.Item2, currentTickSignalTime, runningTrigger.Item3, runningTrigger.Item4)
                                            Next
                                        End If
                                    End If

                                    _canceller.Token.ThrowIfCancellationRequested()
                                    Dim entryTriggers As List(Of Tuple(Of Trade, Payload)) = Nothing
                                    entryTriggers = Await stockStrategyRule.IsTriggerReceivedForPlaceOrderAsync(currentTickSignalTime).ConfigureAwait(False)
                                    If entryTriggers IsNot Nothing AndAlso entryTriggers.Count > 0 Then
                                        For Each runningTrigger In entryTriggers
                                            _canceller.Token.ThrowIfCancellationRequested()
                                            PlaceOrder(runningTrigger.Item1, runningTrigger.Item2, currentTickSignalTime)
                                        Next
                                    End If
                                Next

                                startMinute = startMinute.Add(TimeSpan.FromMinutes(Me.SignalTimeFrame))
                            End While
                        End If
                    End If
                    PopulateDayWiseActiveTradeCount(tradeCheckingDate)
                    tradeCheckingDate = tradeCheckingDate.AddDays(1)
                End While

                'Serialization
                SerializeAllCollections(tradesFileName, capitalFileName, activeTradeFileName)

                PrintArrayToExcel(filename, tradesFileName, capitalFileName, activeTradeFileName, _canceller)
            End If
        End Function

#Region "Stock Selection"
        Private Function GetStockData(ByVal tradingDate As Date) As List(Of StockDetails)
            Dim ret As List(Of StockDetails) = Nothing
            If Me.StockFileName IsNot Nothing Then
                Dim dt As DataTable = Nothing
                Using csvHelper As New Utilities.DAL.CSVHelper(Me.StockFileName, ",", _canceller)
                    dt = csvHelper.GetDataTableFromCSV(1)
                End Using
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    For Each runningRow As DataRow In dt.Rows
                        Dim rowDate As Date = runningRow.Item("Date")
                        If rowDate.Date = tradingDate.Date Then
                            Dim tradingSymbol As String = runningRow.Item("Trading Symbol")
                            Dim instrumentName As String = Nothing
                            If tradingSymbol.Contains("FUT") Then
                                instrumentName = tradingSymbol.Remove(tradingSymbol.Count - 8)
                            Else
                                instrumentName = tradingSymbol
                            End If
                            Dim lotSize As Integer = runningRow.Item("Lot Size")
                            Dim targetLeftPercentage As Decimal = runningRow.Item("Target Left %")

                            If targetLeftPercentage >= 75 Then
                                Dim detailsOfStock As StockDetails = New StockDetails With
                                        {.TradingSymbol = tradingSymbol,
                                         .LotSize = lotSize}
                                'Dim detailsOfStock As StockDetails = New StockDetails With
                                '            {.TradingSymbol = "ICICIBANK",
                                '             .LotSize = 1375}

                                If ret Is Nothing Then ret = New List(Of StockDetails)
                                ret.Add(detailsOfStock)
                            End If
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
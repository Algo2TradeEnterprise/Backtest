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
                'Dim folderpath As String = Path.Combine(My.Application.Info.DirectoryPath, "BackTest Output")
                'Dim files() = Directory.GetFiles(folderpath, "*.xlsx")
                'For Each file In files
                '    If file.ToUpper.Contains(filename.ToUpper) Then
                '        Exit Function
                '    End If
                'Next
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
                    Dim stockList As List(Of StockDetails) = GetStockData(tradeCheckingDate)
                    _canceller.Token.ThrowIfCancellationRequested()
                    If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                        OnHeartbeat("Adding Running stocks")
                        If Me.TradesTaken IsNot Nothing AndAlso Me.TradesTaken.Count > 0 Then
                            For Each runningDate In Me.TradesTaken.Keys
                                For Each runningStock In Me.TradesTaken(runningDate).Keys
                                    Dim availableStock As StockDetails = stockList.Find(Function(x)
                                                                                            Return x.TradingSymbol.ToUpper = runningStock.ToUpper
                                                                                        End Function)
                                    If availableStock Is Nothing Then
                                        Dim lastEntryTrade As Trade = GetLastEntryTradeOfTheStock(runningStock, tradeCheckingDate, Trade.TypeOfTrade.CNC)
                                        If lastEntryTrade IsNot Nothing AndAlso (lastEntryTrade.ExitRemark Is Nothing OrElse
                                            Not lastEntryTrade.ExitRemark.Contains("Target")) Then
                                            availableStock = New StockDetails With {
                                                .TradingSymbol = runningStock,
                                                .LotSize = lastEntryTrade.LotSize
                                            }

                                            stockList.Add(availableStock)
                                        End If
                                    End If
                                Next
                            Next
                        End If

                        Dim checkForEntryExit As Boolean = False
                        Dim stocksRuleData As Dictionary(Of String, StrategyRule) = Nothing

                        'First lets build the payload for all the stocks
                        Dim stockCount As Integer = 0
                        Dim eligibleStockCount As Integer = 0
                        Dim nextTradingDay As Date = Cmn.GetNexTradingDay(Me.DatabaseTable, tradeCheckingDate)
                        'If nextTradingDay = Date.MinValue Then
                        '    nextTradingDay = Now.Date
                        'End If
                        For Each runningPair In stockList
                            _canceller.Token.ThrowIfCancellationRequested()
                            stockCount += 1

                            Dim XDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                            Dim currentDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                            If Me.DataSource = SourceOfData.Database Then
                                XDayOneMinutePayload = Cmn.GetRawPayloadForSpecificTradingSymbol(Me.DatabaseTable, runningPair.TradingSymbol, tradeCheckingDate.AddDays(-30), tradeCheckingDate)
                            ElseIf Me.DataSource = SourceOfData.Live Then
                                Throw New NotImplementedException
                            End If

                            _canceller.Token.ThrowIfCancellationRequested()
                            'Now transfer only the current date payload into the workable payload (this will be used for the main loop and checking if the date is a valid date)
                            If XDayOneMinutePayload IsNot Nothing AndAlso XDayOneMinutePayload.Count > 0 Then
                                OnHeartbeat(String.Format("Processing for {0} on {1}. Stock Counter: [ {2}/{3} ]", runningPair, tradeCheckingDate.ToShortDateString, stockCount, stockList.Count))
                                For Each runningPayload In XDayOneMinutePayload.Keys
                                    _canceller.Token.ThrowIfCancellationRequested()
                                    If runningPayload.Date = tradeCheckingDate.Date Then
                                        If currentDayOneMinutePayload Is Nothing Then currentDayOneMinutePayload = New Dictionary(Of Date, Payload)
                                        currentDayOneMinutePayload.Add(runningPayload, XDayOneMinutePayload(runningPayload))
                                    End If
                                Next
                                'Add all these payloads into the stock collections
                                If currentDayOneMinutePayload IsNot Nothing AndAlso currentDayOneMinutePayload.Count > 0 Then
                                    If stocksRuleData Is Nothing Then stocksRuleData = New Dictionary(Of String, StrategyRule)
                                    Dim stockRule As StrategyRule = Nothing

                                    Select Case Me.RuleNumber
                                        Case 0
                                            stockRule = New HourlyRainbowStrategyRule(tradeCheckingDate, nextTradingDay, runningPair.TradingSymbol, runningPair.LotSize, Me.RuleEntityData, Me, _canceller, XDayOneMinutePayload)
                                        Case Else
                                            Throw New NotImplementedException
                                    End Select

                                    AddHandler stockRule.Heartbeat, AddressOf OnHeartbeat
                                    stockRule.CompletePreProcessing()
                                    stocksRuleData.Add(runningPair.TradingSymbol, stockRule)

                                    checkForEntryExit = True
                                Else
                                    If Cmn.IsTradingDay(tradeCheckingDate) Then
                                        Throw New NotImplementedException(String.Format("Data unavailable for {0}, {1}", runningPair.TradingSymbol, tradeCheckingDate.ToString("dd-MMM-yyyy")))
                                    End If
                                End If
                            Else
                                If Cmn.IsTradingDay(tradeCheckingDate) Then
                                    Throw New NotImplementedException(String.Format("Data unavailable for {0}, {1}", runningPair.TradingSymbol, tradeCheckingDate.ToString("dd-MMM-yyyy")))
                                End If
                            End If
                        Next
                        '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                        If checkForEntryExit Then
                            OnHeartbeat(String.Format("Checking Trade on {0}", tradeCheckingDate.ToShortDateString))
                            _canceller.Token.ThrowIfCancellationRequested()
                            Dim tradeStartTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, Me.TradeStartTime.Hours, Me.TradeStartTime.Minutes, Me.TradeStartTime.Seconds)
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
                                        For Each runningStock In stockList
                                            _canceller.Token.ThrowIfCancellationRequested()

                                            Dim stockStrategyRule As StrategyRule = stocksRuleData(runningStock.TradingSymbol)

                                            _canceller.Token.ThrowIfCancellationRequested()
                                            Dim potentialRuleExitTrades As List(Of Trade) = GetSpecificTrades(runningStock.TradingSymbol, tradeCheckingDate, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
                                            If potentialRuleExitTrades IsNot Nothing AndAlso potentialRuleExitTrades.Count > 0 Then
                                                Await stockStrategyRule.IsTriggerReceivedForExitOrderAsync(potentialTickSignalTime, potentialRuleExitTrades).ConfigureAwait(False)
                                            End If

                                            If potentialTickSignalTime = GetCurrentXMinuteCandleTime(potentialTickSignalTime) OrElse
                                                potentialTickSignalTime = tradeStartTime Then
                                                _canceller.Token.ThrowIfCancellationRequested()
                                                Await stockStrategyRule.IsTriggerReceivedForPlaceOrderAsync(potentialTickSignalTime).ConfigureAwait(False)
                                            End If
                                        Next
                                    End If
                                    startSecond = startSecond.Add(TimeSpan.FromSeconds(1))
                                End While   'Second Loop
                                startMinute = startMinute.Add(TimeSpan.FromMinutes(Me.SignalTimeFrame))
                            End While   'Minute Loop
                        End If
                    End If
                    'SetOverallDrawUpDrawDownForTheDay(tradeCheckingDate)
                    totalPL += TotalPLAfterBrokerage(tradeCheckingDate)
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

                            Dim detailsOfStock As StockDetails = New StockDetails With
                                        {.TradingSymbol = tradingSymbol,
                                         .LotSize = lotSize}
                            'Dim detailsOfStock As StockDetails = New StockDetails With
                            '            {.TradingSymbol = "ICICIBANK",
                            '             .LotSize = 1375}

                            If ret Is Nothing Then ret = New List(Of StockDetails)
                            ret.Add(detailsOfStock)
                            If ret.Count >= Me.NumberOfTradeableStockPerDay Then Exit For
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
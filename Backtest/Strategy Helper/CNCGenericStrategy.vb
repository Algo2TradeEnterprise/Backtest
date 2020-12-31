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
                    Dim pairList As List(Of PairDetails) = GetStockData(tradeCheckingDate)
                    _canceller.Token.ThrowIfCancellationRequested()
                    If pairList IsNot Nothing AndAlso pairList.Count > 0 Then
                        Dim checkForEntryExit As Boolean = False
                        Dim stocksRuleData As Dictionary(Of String, PairStrategyRule) = Nothing

                        'First lets build the payload for all the stocks
                        Dim stockCount As Integer = 0
                        Dim eligibleStockCount As Integer = 0
                        Dim commonNextTradingDay As Date = Cmn.GetNexTradingDay(Me.DatabaseTable, tradeCheckingDate)
                        For Each runningPair In pairList
                            _canceller.Token.ThrowIfCancellationRequested()
                            stockCount += 1

                            Dim XDayOneMinutePayload1 As Dictionary(Of Date, Payload) = Nothing
                            Dim XDayOneMinutePayload2 As Dictionary(Of Date, Payload) = Nothing
                            Dim currentDayOneMinutePayload1 As Dictionary(Of Date, Payload) = Nothing
                            Dim currentDayOneMinutePayload2 As Dictionary(Of Date, Payload) = Nothing
                            If Me.DataSource = SourceOfData.Database Then
                                XDayOneMinutePayload1 = Cmn.GetRawPayloadForSpecificTradingSymbol(Me.DatabaseTable, runningPair.TradingSymbol1, tradeCheckingDate.AddDays(-30), tradeCheckingDate)
                                XDayOneMinutePayload2 = Cmn.GetRawPayloadForSpecificTradingSymbol(Me.DatabaseTable, runningPair.TradingSymbol2, tradeCheckingDate.AddDays(-30), tradeCheckingDate)
                            ElseIf Me.DataSource = SourceOfData.Live Then
                                Throw New NotImplementedException
                            End If

                            _canceller.Token.ThrowIfCancellationRequested()
                            'Now transfer only the current date payload into the workable payload (this will be used for the main loop and checking if the date is a valid date)
                            If XDayOneMinutePayload1 IsNot Nothing AndAlso XDayOneMinutePayload1.Count > 0 AndAlso
                                XDayOneMinutePayload2 IsNot Nothing AndAlso XDayOneMinutePayload2.Count > 0 Then
                                OnHeartbeat(String.Format("Processing for {0} on {1}. Stock Counter: [ {2}/{3} ]", runningPair, tradeCheckingDate.ToShortDateString, stockCount, pairList.Count))
                                For Each runningPayload In XDayOneMinutePayload1.Keys
                                    _canceller.Token.ThrowIfCancellationRequested()
                                    If runningPayload.Date = tradeCheckingDate.Date Then
                                        If currentDayOneMinutePayload1 Is Nothing Then currentDayOneMinutePayload1 = New Dictionary(Of Date, Payload)
                                        currentDayOneMinutePayload1.Add(runningPayload, XDayOneMinutePayload1(runningPayload))
                                    End If
                                Next
                                For Each runningPayload In XDayOneMinutePayload2.Keys
                                    _canceller.Token.ThrowIfCancellationRequested()
                                    If runningPayload.Date = tradeCheckingDate.Date Then
                                        If currentDayOneMinutePayload2 Is Nothing Then currentDayOneMinutePayload2 = New Dictionary(Of Date, Payload)
                                        currentDayOneMinutePayload2.Add(runningPayload, XDayOneMinutePayload2(runningPayload))
                                    End If
                                Next
                                'Add all these payloads into the stock collections
                                If currentDayOneMinutePayload1 IsNot Nothing AndAlso currentDayOneMinutePayload1.Count > 0 AndAlso
                                    currentDayOneMinutePayload2 IsNot Nothing AndAlso currentDayOneMinutePayload2.Count > 0 Then
                                    If stocksRuleData Is Nothing Then stocksRuleData = New Dictionary(Of String, PairStrategyRule)
                                    Dim stockRule As PairStrategyRule = Nothing

                                    stockRule = New OutsidexSDStrategyRule(runningPair.PairName, tradeCheckingDate, commonNextTradingDay, runningPair.TradingSymbol1, runningPair.LotSize1, runningPair.TradingSymbol2, runningPair.LotSize2, Me.RuleEntityData, Me, _canceller, XDayOneMinutePayload1, XDayOneMinutePayload2)

                                    AddHandler stockRule.Heartbeat, AddressOf OnHeartbeat
                                    stockRule.CompletePreProcessing()
                                    stocksRuleData.Add(runningPair.PairName, stockRule)

                                    checkForEntryExit = True
                                Else
                                    If Cmn.IsTradingDay(tradeCheckingDate) Then
                                        Throw New NotImplementedException(String.Format("Data unavailable for {0}, {1}", runningPair.PairName, tradeCheckingDate.ToString("dd-MMM-yyyy")))
                                    End If
                                End If
                            Else
                                If Cmn.IsTradingDay(tradeCheckingDate) Then
                                    Throw New NotImplementedException(String.Format("Data unavailable for {0}, {1}", runningPair.PairName, tradeCheckingDate.ToString("dd-MMM-yyyy")))
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
                                        For Each runningPair In pairList
                                            _canceller.Token.ThrowIfCancellationRequested()

                                            Dim stockStrategyRule As PairStrategyRule = stocksRuleData(runningPair.PairName)

                                            _canceller.Token.ThrowIfCancellationRequested()
                                            Dim potentialRuleExitTrades As List(Of Trade) = GetSpecificTrades(runningPair.PairName, tradeCheckingDate, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
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
        Private Function GetStockData(ByVal tradingDate As Date) As List(Of PairDetails)
            Dim ret As List(Of PairDetails) = Nothing

            Dim detailsOfStock As PairDetails = New PairDetails With
                                    {.PairName = "HDFC~HDFCBANK",
                                     .TradingSymbol1 = "HDFC",
                                     .LotSize1 = 300,
                                     .TradingSymbol2 = "HDFCBANK",
                                     .LotSize2 = 550}

            ret = New List(Of PairDetails)
            ret.Add(detailsOfStock)

            'If Me.StockFileName IsNot Nothing Then
            '    Dim dt As DataTable = Nothing
            '    Using csvHelper As New Utilities.DAL.CSVHelper(Me.StockFileName, ",", _canceller)
            '        dt = csvHelper.GetDataTableFromCSV(1)
            '    End Using
            '    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            '        For Each runningRow As DataRow In dt.Rows
            '            If ret Is Nothing Then ret = New List(Of StockDetails)
            '            Dim tradingSymbol As String = runningRow.Item(0)
            '            Dim instrumentName As String = Nothing
            '            If tradingSymbol.Contains("FUT") Then
            '                instrumentName = tradingSymbol.Remove(tradingSymbol.Count - 8)
            '            Else
            '                instrumentName = tradingSymbol
            '            End If
            '            Dim lotSize As Integer = runningRow.Item(1)
            '            Dim xlStockType As String = runningRow.Item(2)
            '            Dim controller As String = runningRow.Item(3)
            '            Dim stockTyp As Trade.TypeOfStock = Trade.TypeOfStock.None
            '            Select Case xlStockType
            '                Case "Cash"
            '                    stockTyp = Trade.TypeOfStock.Cash
            '                Case "Future"
            '                    stockTyp = Trade.TypeOfStock.Futures
            '                Case Else
            '                    Throw New NotImplementedException
            '            End Select

            '            Dim detailsOfStock As StockDetails = New StockDetails With
            '                        {.StockName = instrumentName,
            '                         .TradingSymbol = tradingSymbol,
            '                         .LotSize = lotSize,
            '                         .EligibleToTakeTrade = True,
            '                         .StockType = stockTyp,
            '                         .Supporting1 = If(controller.ToUpper = "TRUE", 1, 0)}

            '            ret.Add(detailsOfStock)
            '        Next
            '    End If
            'End If
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
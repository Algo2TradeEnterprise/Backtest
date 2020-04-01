Imports System.IO
Imports Utilities.DAL
Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Strings
Imports Backtest.StrategyHelper

Public Class DataFetcher

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

    Private ReadOnly _cts As CancellationTokenSource
    Private ReadOnly _serverName As String
    Private ReadOnly _uniqueInstrumentList As List(Of String)
    Private ReadOnly _startDate As Date
    Private ReadOnly _endDate As Date
    Private ReadOnly _stockType As Trade.TypeOfStock
    Private ReadOnly _directoryName As String

    Public Sub New(ByVal canceller As CancellationTokenSource,
                   ByVal serverName As String,
                   ByVal uniqueInstrumentList As List(Of String),
                   ByVal startDate As Date,
                   ByVal endDate As Date,
                   ByVal stockType As Trade.TypeOfStock,
                   ByVal strategyName As String)
        _cts = canceller
        _serverName = serverName
        _uniqueInstrumentList = uniqueInstrumentList
        _startDate = startDate
        _endDate = endDate
        _stockType = stockType
        _directoryName = Path.Combine(My.Application.Info.DirectoryPath, String.Format("{0} CANDLE DATA", strategyName.ToUpper))
    End Sub

    Public Async Function GetCandleData(ByVal tradingSymbol As String, ByVal startDate As Date, ByVal endDate As Date) As Task(Of Dictionary(Of Date, Payload))
        Dim ret As Dictionary(Of Date, Payload) = Nothing
        If Not Directory.Exists(_directoryName) Then
            Directory.CreateDirectory(_directoryName)
            _cts.Token.ThrowIfCancellationRequested()
            Dim success As Boolean = Await GetAllData().ConfigureAwait(False)
            If Not success Then Throw New ApplicationException("Unable to any data from database")
        End If
        If Directory.Exists(_directoryName) Then
            Dim filename As String = Path.Combine(_directoryName, String.Format("{0}.xml", tradingSymbol))
            If File.Exists(filename) Then
                Dim ds As DataSet = New DataSet()
                _cts.Token.ThrowIfCancellationRequested()
                ds.ReadXml(filename, XmlReadMode.InferSchema)
                _cts.Token.ThrowIfCancellationRequested()
                If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                    Dim tempdt As DataTable = ds.Tables(0)
                    If tempdt IsNot Nothing AndAlso tempdt.Rows.Count > 0 Then
                        _cts.Token.ThrowIfCancellationRequested()
                        Dim rows As DataRow() = tempdt.Select(String.Format("SnapshotDateTime>=#{0}# AND SnapshotDateTime<=#{1}#",
                                                                            startDate.Date.ToString("yyyy-MM-dd HH:mm:ss"),
                                                                            endDate.AddDays(1).Date.ToString("yyyy-MM-dd HH:mm:ss")))
                        _cts.Token.ThrowIfCancellationRequested()
                        If rows IsNot Nothing Then
                            Dim dt As DataTable = New DataTable()
                            dt.Columns.Add("Open")
                            dt.Columns.Add("Low")
                            dt.Columns.Add("High")
                            dt.Columns.Add("Close")
                            dt.Columns.Add("Volume")
                            dt.Columns.Add("SnapshotDateTime")
                            dt.Columns.Add("TradingSymbol")
                            For Each runningRow In rows
                                _cts.Token.ThrowIfCancellationRequested()
                                dt.Rows.Add(runningRow.Item("Open"),
                                            runningRow.Item("Low"),
                                            runningRow.Item("High"),
                                            runningRow.Item("Close"),
                                            runningRow.Item("Volume"),
                                            runningRow.Item("SnapshotDateTime"),
                                            runningRow.Item("TradingSymbol"))
                            Next
                            _cts.Token.ThrowIfCancellationRequested()
                            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                ret = Common.ConvertDataTableToPayload(dt, 0, 1, 2, 3, 4, 5, 6)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Async Function GetAllData() As Task(Of Boolean)
        Dim ret As Boolean = False
        If _uniqueInstrumentList IsNot Nothing AndAlso _uniqueInstrumentList.Count > 0 Then
            _cts.Token.ThrowIfCancellationRequested()
            Using sqlHlpr As New MySQLDBHelper(_serverName, "local_stock", "3306", "rio", "speech123", _cts)
                AddHandler sqlHlpr.Heartbeat, AddressOf OnHeartbeat
                AddHandler sqlHlpr.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
                AddHandler sqlHlpr.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
                AddHandler sqlHlpr.WaitingFor, AddressOf OnWaitingFor

                Dim tableName As String = Nothing
                Select Case _stockType
                    Case Trade.TypeOfStock.Cash
                        tableName = "intraday_prices_cash"
                    Case Trade.TypeOfStock.Commodity
                        tableName = "intraday_prices_commodity"
                    Case Trade.TypeOfStock.Currency
                        tableName = "intraday_prices_currency"
                    Case Trade.TypeOfStock.Futures
                        tableName = "intraday_prices_futures"
                    Case Else
                        Throw New NotImplementedException
                End Select

                Dim queryString As String = Nothing
                For Each runningStock In _uniqueInstrumentList
                    _cts.Token.ThrowIfCancellationRequested()
                    queryString = String.Format("{0}SELECT `Open`,`Low`,`High`,`Close`,`Volume`,`SnapshotDateTime`,`TradingSymbol` FROM `{1}` WHERE `TradingSymbol`='{2}' AND `SnapshotDate`>='{3}' AND `SnapshotDate`<='{4}';",
                                                queryString, tableName, runningStock, _startDate.ToString("yyyy-MM-dd"), _endDate.ToString("yyyy-MM-dd"))
                Next

                _cts.Token.ThrowIfCancellationRequested()
                Dim ds As DataSet = Await sqlHlpr.RunSelectSetAsync(queryString).ConfigureAwait(False)
                _cts.Token.ThrowIfCancellationRequested()
                If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                    For tableNumber As Integer = 0 To ds.Tables.Count - 1
                        _cts.Token.ThrowIfCancellationRequested()
                        Dim dt As DataTable = ds.Tables(tableNumber)
                        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            Dim instrumentName As String = dt.Rows(0).Item("TradingSymbol")
                            Dim filename As String = Path.Combine(_directoryName, String.Format("{0}.xml", instrumentName))
                            _cts.Token.ThrowIfCancellationRequested()
                            dt.WriteXml(filename)
                            _cts.Token.ThrowIfCancellationRequested()
                            ret = True
                        End If
                    Next
                End If
            End Using
        End If
        Return ret
    End Function
End Class

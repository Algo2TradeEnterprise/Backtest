Imports Utilities.Network
Imports Utilities.DAL
Imports System.Threading
Imports System.IO

Namespace StrategyHelper
    Public Class SectoralStockFetcher
        Implements IDisposable

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

#Region "Enum"
        Public Enum SectorType
            NIFTYAUTO = 1
            NIFTYBANK
            NIFTYCONSUMER
            NIFTYFINSERV
            NIFTYFMCG
            NIFTYHEALTHCARE
            NIFTYIT
            NIFTYMEDIA
            NIFTYMETAL
            NIFTYOILGAS
            NIFTYPHARMA
            NIFTYPVTBANK
            NIFTYPSUBANK
            NIFTYREALTY
        End Enum
#End Region

        Private ReadOnly _cts As CancellationTokenSource
        Public Sub New(ByVal canceller As CancellationTokenSource)
            _cts = canceller
        End Sub

        Private Function GetSectorURL(ByVal typeOfSector As SectorType) As String
            Dim ret As String = Nothing
            Select Case typeOfSector
                Case SectorType.NIFTYAUTO
                    ret = "https://www1.nseindia.com/content/indices/ind_niftyautolist.csv"
                Case SectorType.NIFTYBANK
                    ret = "https://www1.nseindia.com/content/indices/ind_niftybanklist.csv"
                Case SectorType.NIFTYCONSUMER
                    ret = "https://www1.nseindia.com/content/indices/ind_niftyconsumerdurableslist.csv"
                Case SectorType.NIFTYFINSERV
                    ret = "https://www1.nseindia.com/content/indices/ind_niftyfinancelist.csv"
                Case SectorType.NIFTYFMCG
                    ret = "https://www1.nseindia.com/content/indices/ind_niftyfmcglist.csv"
                Case SectorType.NIFTYHEALTHCARE
                    ret = "https://www1.nseindia.com/content/indices/ind_niftyhealthcarelist.csv"
                Case SectorType.NIFTYIT
                    ret = "https://www1.nseindia.com/content/indices/ind_niftyitlist.csv"
                Case SectorType.NIFTYMEDIA
                    ret = "https://www1.nseindia.com/content/indices/ind_niftymedialist.csv"
                Case SectorType.NIFTYMETAL
                    ret = "https://www1.nseindia.com/content/indices/ind_niftymetallist.csv"
                Case SectorType.NIFTYOILGAS
                    ret = "https://www1.nseindia.com/content/indices/ind_niftyoilgaslist.csv"
                Case SectorType.NIFTYPHARMA
                    ret = "https://www1.nseindia.com/content/indices/ind_niftypharmalist.csv"
                Case SectorType.NIFTYPVTBANK
                    ret = "https://www1.nseindia.com/content/indices/ind_nifty_privatebanklist.csv"
                Case SectorType.NIFTYPSUBANK
                    ret = "https://www1.nseindia.com/content/indices/ind_niftypsubanklist.csv"
                Case SectorType.NIFTYREALTY
                    ret = "https://www1.nseindia.com/content/indices/ind_niftyrealtylist.csv"
                Case Else
                    Throw New NotImplementedException
            End Select
            Return ret
        End Function

        Private Async Function GetSectorStockFileAsync(ByVal typeOfSector As SectorType, ByVal filename As String) As Task(Of Boolean)
            Dim ret As Boolean = False
            Using browser As New HttpBrowser(Nothing, Net.DecompressionMethods.GZip, TimeSpan.FromSeconds(30), _cts)
                AddHandler browser.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
                AddHandler browser.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
                AddHandler browser.WaitingFor, AddressOf OnWaitingFor
                AddHandler browser.Heartbeat, AddressOf OnHeartbeat

                browser.KeepAlive = True
                Dim headersToBeSent As New Dictionary(Of String, String)
                headersToBeSent.Add("Host", "www1.nseindia.com")
                headersToBeSent.Add("Upgrade-Insecure-Requests", "1")
                headersToBeSent.Add("Sec-Fetch-Mode", "navigate")
                headersToBeSent.Add("Sec-Fetch-Site", "none")

                Dim targetURL As String = GetSectorURL(typeOfSector)
                If targetURL IsNot Nothing Then
                    ret = Await browser.GetFileAsync(targetURL, filename, False, headersToBeSent).ConfigureAwait(False)
                End If
            End Using
            Return ret
        End Function

        Public Async Function GetSectoralStocklist(ByVal typeOfSector As SectorType) As Task(Of List(Of String))
            Dim ret As List(Of String) = Nothing
            Dim filename As String = Path.Combine(My.Application.Info.DirectoryPath, String.Format("Sectoral Index {0}.csv", typeOfSector.ToString))
            Dim fileAvailable As Boolean = Await GetSectorStockFileAsync(typeOfSector, filename).ConfigureAwait(False)
            If fileAvailable AndAlso File.Exists(filename) Then
                OnHeartbeat("Reading stock file")
                Dim dt As DataTable = Nothing
                Using csv As New CSVHelper(filename, ",", _cts)
                    'AddHandler csv.Heartbeat, AddressOf OnHeartbeat
                    dt = csv.GetDataTableFromCSV(0)
                End Using
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        _cts.Token.ThrowIfCancellationRequested()
                        If ret Is Nothing Then ret = New List(Of String)
                        ret.Add(dt.Rows(i).Item("Symbol").ToString.ToUpper)
                    Next
                End If
            End If
            If File.Exists(filename) Then File.Delete(filename)
            Return ret
        End Function

        Public Function GetAllSectorList() As List(Of SectorType)
            Dim ret As List(Of SectorType) = Nothing
            For counter As Integer = 1 To 100
                _cts.Token.ThrowIfCancellationRequested()
                Dim sector As SectorType = counter
                If Not IsNumeric(sector.ToString) Then
                    If ret Is Nothing Then ret = New List(Of SectorType)
                    ret.Add(sector)
                Else
                    Exit For
                End If
            Next
            Return ret
        End Function

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
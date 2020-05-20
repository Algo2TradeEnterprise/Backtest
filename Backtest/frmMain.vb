Imports Algo2TradeBLL
Imports System.IO
Imports System.Threading
Imports Backtest.StrategyHelper
Public Class frmMain

#Region "Common Delegates"
    Delegate Sub SetObjectEnableDisable_Delegate(ByVal [obj] As Object, ByVal [value] As Boolean)
    Public Sub SetObjectEnableDisable_ThreadSafe(ByVal [obj] As Object, ByVal [value] As Boolean)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [obj].InvokeRequired Then
            Dim MyDelegate As New SetObjectEnableDisable_Delegate(AddressOf SetObjectEnableDisable_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[obj], [value]})
        Else
            [obj].Enabled = [value]
        End If
    End Sub

    Delegate Sub SetObjectVisible_Delegate(ByVal [obj] As Object, ByVal [value] As Boolean)
    Public Sub SetObjectVisible_ThreadSafe(ByVal [obj] As Object, ByVal [value] As Boolean)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [obj].InvokeRequired Then
            Dim MyDelegate As New SetObjectVisible_Delegate(AddressOf SetObjectVisible_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[obj], [value]})
        Else
            [obj].Visible = [value]
        End If
    End Sub

    Delegate Sub SetLabelText_Delegate(ByVal [label] As Label, ByVal [text] As String)
    Public Sub SetLabelText_ThreadSafe(ByVal [label] As Label, ByVal [text] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New SetLabelText_Delegate(AddressOf SetLabelText_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[label], [text]})
        Else
            [label].Text = [text]
        End If
    End Sub

    Delegate Function GetLabelText_Delegate(ByVal [label] As Label) As String
    Public Function GetLabelText_ThreadSafe(ByVal [label] As Label) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New GetLabelText_Delegate(AddressOf GetLabelText_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[label]})
        Else
            Return [label].Text
        End If
    End Function

    Delegate Sub SetLabelTag_Delegate(ByVal [label] As Label, ByVal [tag] As String)
    Public Sub SetLabelTag_ThreadSafe(ByVal [label] As Label, ByVal [tag] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New SetLabelTag_Delegate(AddressOf SetLabelTag_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[label], [tag]})
        Else
            [label].Tag = [tag]
        End If
    End Sub

    Delegate Function GetLabelTag_Delegate(ByVal [label] As Label) As String
    Public Function GetLabelTag_ThreadSafe(ByVal [label] As Label) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New GetLabelTag_Delegate(AddressOf GetLabelTag_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[label]})
        Else
            Return [label].Tag
        End If
    End Function
    Delegate Sub SetToolStripLabel_Delegate(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripStatusLabel, ByVal [text] As String)
    Public Sub SetToolStripLabel_ThreadSafe(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripStatusLabel, ByVal [text] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [toolStrip].InvokeRequired Then
            Dim MyDelegate As New SetToolStripLabel_Delegate(AddressOf SetToolStripLabel_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[toolStrip], [label], [text]})
        Else
            [label].Text = [text]
        End If
    End Sub

    Delegate Function GetToolStripLabel_Delegate(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripLabel) As String
    Public Function GetToolStripLabel_ThreadSafe(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripLabel) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [toolStrip].InvokeRequired Then
            Dim MyDelegate As New GetToolStripLabel_Delegate(AddressOf GetToolStripLabel_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[toolStrip], [label]})
        Else
            Return [label].Text
        End If
    End Function

    Delegate Function GetDateTimePickerValue_Delegate(ByVal [dateTimePicker] As DateTimePicker) As Date
    Public Function GetDateTimePickerValue_ThreadSafe(ByVal [dateTimePicker] As DateTimePicker) As Date
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [dateTimePicker].InvokeRequired Then
            Dim MyDelegate As New GetDateTimePickerValue_Delegate(AddressOf GetDateTimePickerValue_ThreadSafe)
            Return Me.Invoke(MyDelegate, New DateTimePicker() {[dateTimePicker]})
        Else
            Return [dateTimePicker].Value
        End If
    End Function

    Delegate Function GetNumericUpDownValue_Delegate(ByVal [numericUpDown] As NumericUpDown) As Integer
    Public Function GetNumericUpDownValue_ThreadSafe(ByVal [numericUpDown] As NumericUpDown) As Integer
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [numericUpDown].InvokeRequired Then
            Dim MyDelegate As New GetNumericUpDownValue_Delegate(AddressOf GetNumericUpDownValue_ThreadSafe)
            Return Me.Invoke(MyDelegate, New NumericUpDown() {[numericUpDown]})
        Else
            Return [numericUpDown].Value
        End If
    End Function

    Delegate Function GetComboBoxIndex_Delegate(ByVal [combobox] As ComboBox) As Integer
    Public Function GetComboBoxIndex_ThreadSafe(ByVal [combobox] As ComboBox) As Integer
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [combobox].InvokeRequired Then
            Dim MyDelegate As New GetComboBoxIndex_Delegate(AddressOf GetComboBoxIndex_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[combobox]})
        Else
            Return [combobox].SelectedIndex
        End If
    End Function

    Delegate Function GetComboBoxItem_Delegate(ByVal [ComboBox] As ComboBox) As String
    Public Function GetComboBoxItem_ThreadSafe(ByVal [ComboBox] As ComboBox) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [ComboBox].InvokeRequired Then
            Dim MyDelegate As New GetComboBoxItem_Delegate(AddressOf GetComboBoxItem_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[ComboBox]})
        Else
            Return [ComboBox].SelectedItem.ToString
        End If
    End Function

    Delegate Function GetTextBoxText_Delegate(ByVal [textBox] As TextBox) As String
    Public Function GetTextBoxText_ThreadSafe(ByVal [textBox] As TextBox) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [textBox].InvokeRequired Then
            Dim MyDelegate As New GetTextBoxText_Delegate(AddressOf GetTextBoxText_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[textBox]})
        Else
            Return [textBox].Text
        End If
    End Function

    Delegate Function GetCheckBoxChecked_Delegate(ByVal [checkBox] As CheckBox) As Boolean
    Public Function GetCheckBoxChecked_ThreadSafe(ByVal [checkBox] As CheckBox) As Boolean
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [checkBox].InvokeRequired Then
            Dim MyDelegate As New GetCheckBoxChecked_Delegate(AddressOf GetCheckBoxChecked_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[checkBox]})
        Else
            Return [checkBox].Checked
        End If
    End Function

    Delegate Function GetRadioButtonChecked_Delegate(ByVal [radioButton] As RadioButton) As Boolean
    Public Function GetRadioButtonChecked_ThreadSafe(ByVal [radioButton] As RadioButton) As Boolean
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [radioButton].InvokeRequired Then
            Dim MyDelegate As New GetRadioButtonChecked_Delegate(AddressOf GetRadioButtonChecked_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[radioButton]})
        Else
            Return [radioButton].Checked
        End If
    End Function

    Delegate Sub SetDatagridBindDatatable_Delegate(ByVal [datagrid] As DataGridView, ByVal [table] As DataTable)
    Public Sub SetDatagridBindDatatable_ThreadSafe(ByVal [datagrid] As DataGridView, ByVal [table] As DataTable)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [datagrid].InvokeRequired Then
            Dim MyDelegate As New SetDatagridBindDatatable_Delegate(AddressOf SetDatagridBindDatatable_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[datagrid], [table]})
        Else
            [datagrid].DataSource = [table]
            [datagrid].Refresh()
        End If
    End Sub
#End Region

#Region "Event Handlers"
    Private Sub OnHeartbeat(message As String)
        SetLabelText_ThreadSafe(lblProgress, message)
    End Sub
    Private Sub OnHeartbeatMain(message As String)
        SetLabelText_ThreadSafe(lblProgress, message)
    End Sub
    Private Sub OnDocumentDownloadComplete()
        'OnHeartbeat("Document download compelete")
    End Sub
    Private Sub OnDocumentRetryStatus(currentTry As Integer, totalTries As Integer)
        OnHeartbeat(String.Format("Try #{0}/{1}: Connecting...", currentTry, totalTries))
    End Sub
    Public Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        OnHeartbeat(String.Format("{0}, waiting {1}/{2} secs", msg, elapsedSecs, totalSecs))
    End Sub
#End Region

    Private _canceller As CancellationTokenSource

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Settings.Rule <= cmbRule.Items.Count - 1 Then cmbRule.SelectedIndex = My.Settings.Rule
        If My.Settings.StartDate <> Date.MinValue Then dtpckrStartDate.Value = My.Settings.StartDate
        If My.Settings.EndDate <> Date.MinValue Then dtpckrEndDate.Value = My.Settings.EndDate

        rdbMIS.Checked = My.Settings.MIS
        rdbCNCTick.Checked = My.Settings.CNCTick
        rdbCNCCandle.Checked = My.Settings.CNCCandle
        rdbCNCEOD.Checked = My.Settings.CNCEOD

        rdbDatabase.Checked = My.Settings.Database
        rdbLive.Checked = My.Settings.Live
        rdbDatabase_CheckedChanged(sender, e)
        rdbLive_CheckedChanged(sender, e)

        rdbLocalDBConnection.Checked = My.Settings.LocalConnection
        rdbRemoteDBConnection.Checked = My.Settings.RemoteConnection

        SetObjectEnableDisable_ThreadSafe(btnStop, False)
    End Sub

    Private Async Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        _canceller = New CancellationTokenSource
        SetObjectEnableDisable_ThreadSafe(btnStart, False)
        SetObjectEnableDisable_ThreadSafe(btnStop, True)
        My.Settings.Rule = cmbRule.SelectedIndex

        My.Settings.StartDate = dtpckrStartDate.Value
        My.Settings.EndDate = dtpckrEndDate.Value

        My.Settings.MIS = rdbMIS.Checked
        My.Settings.CNCTick = rdbCNCTick.Checked
        My.Settings.CNCEOD = rdbCNCEOD.Checked
        My.Settings.CNCCandle = rdbCNCCandle.Checked

        My.Settings.Database = rdbDatabase.Checked
        My.Settings.Live = rdbLive.Checked

        My.Settings.LocalConnection = rdbLocalDBConnection.Checked
        My.Settings.RemoteConnection = rdbRemoteDBConnection.Checked
        If My.Settings.LocalConnection Then
            My.Settings.ServerName = "localhost"
        Else
            My.Settings.ServerName = "103.57.246.210"
        End If
        My.Settings.Save()

        If rdbMIS.Checked Then
            Await Task.Run(AddressOf ViewDataMISAsync).ConfigureAwait(False)
        ElseIf rdbCNCTick.Checked Then
            Throw New NotImplementedException
        ElseIf rdbCNCEOD.Checked Then
            Throw New NotImplementedException
        ElseIf rdbCNCCandle.Checked Then
            Throw New NotImplementedException
        End If
    End Sub

    Private Async Function ViewDataMISAsync() As Task
        Try
            Dim startDate As Date = GetDateTimePickerValue_ThreadSafe(dtpckrStartDate)
            Dim endDate As Date = GetDateTimePickerValue_ThreadSafe(dtpckrEndDate)
            Dim sourceData As Strategy.SourceOfData = Strategy.SourceOfData.None
            If GetRadioButtonChecked_ThreadSafe(rdbLive) Then
                sourceData = Strategy.SourceOfData.Live
            Else
                sourceData = Strategy.SourceOfData.Database
            End If

            Dim ruleNumber As Integer = GetComboBoxIndex_ThreadSafe(cmbRule)
            Select Case ruleNumber
                Case 0
#Region "Reversal HHLL Breakout"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 10
                            tick = 0.05
                        Case Trade.TypeOfStock.Commodity
                            database = Common.DataBaseTable.Intraday_Commodity
                            margin = 70
                            tick = 1
                        Case Trade.TypeOfStock.Currency
                            database = Common.DataBaseTable.Intraday_Currency
                            margin = 98
                            tick = 0.0025
                        Case Trade.TypeOfStock.Futures
                            database = Common.DataBaseTable.Intraday_Futures
                            margin = 30
                            tick = 0.05
                    End Select

                    Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                            exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                            exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                            tradeStartTime:=TimeSpan.Parse("9:18:00"),
                                                            lastTradeEntryTime:=TimeSpan.Parse("14:29:59"),
                                                            eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                            tickSize:=tick,
                                                            marginMultiplier:=margin,
                                                            timeframe:=1,
                                                            heikenAshiCandle:=False,
                                                            stockType:=stockType,
                                                            databaseTable:=database,
                                                            dataSource:=sourceData,
                                                            initialCapital:=300000,
                                                            usableCapital:=200000,
                                                            minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                            amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Stock.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New ReversalHHLLBreakoutStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .MinimumTargetMultiplier = 2
                                            }

                            .NumberOfTradeableStockPerDay = Integer.MaxValue

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = False
                            .StockMaxProfitPerDay = Decimal.MaxValue
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = True
                            .OverAllProfitPerDay = 15000
                            .OverAllLossPerDay = -5000

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim filename As String = String.Format("Reversal HHLL")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 1
#Region "Fractal Trend Line"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 10
                            tick = 0.05
                        Case Trade.TypeOfStock.Commodity
                            database = Common.DataBaseTable.Intraday_Commodity
                            margin = 70
                            tick = 1
                        Case Trade.TypeOfStock.Currency
                            database = Common.DataBaseTable.Intraday_Currency
                            margin = 98
                            tick = 0.0025
                        Case Trade.TypeOfStock.Futures
                            database = Common.DataBaseTable.Intraday_Futures
                            margin = 30
                            tick = 0.05
                    End Select

                    For trndLnTyp As Integer = 1 To 1
                        Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                    tradeStartTime:=TimeSpan.Parse("9:16:00"),
                                                                    lastTradeEntryTime:=TimeSpan.Parse("14:44:59"),
                                                                    eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                    tickSize:=tick,
                                                                    marginMultiplier:=margin,
                                                                    timeframe:=1,
                                                                    heikenAshiCandle:=False,
                                                                    stockType:=stockType,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=300000,
                                                                    usableCapital:=200000,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                            With backtestStrategy
                                .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "BANKNIFTY.csv")

                                .AllowBothDirectionEntryAtSameTime = False
                                .TrailingStoploss = False
                                .TickBasedStrategy = False
                                .RuleNumber = ruleNumber

                                .RuleEntityData = New FractalTrendLineStrategyRule.StrategyRuleEntities With
                                                {
                                                    .TypeOfTrendLine = trndLnTyp
                                                }

                                .NumberOfTradeableStockPerDay = Integer.MaxValue

                                .NumberOfTradesPerStockPerDay = Integer.MaxValue

                                .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                                .StockMaxLossPercentagePerDay = Decimal.MinValue

                                .ExitOnStockFixedTargetStoploss = False
                                .StockMaxProfitPerDay = Decimal.MaxValue
                                .StockMaxLossPerDay = Decimal.MinValue

                                .ExitOnOverAllFixedTargetStoploss = False
                                .OverAllProfitPerDay = Decimal.MaxValue
                                .OverAllLossPerDay = Decimal.MinValue

                                .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                                .MTMSlab = Math.Abs(.OverAllLossPerDay)
                                .MovementSlab = .MTMSlab / 2
                                .RealtimeTrailingPercentage = 50
                            End With

                            Dim rule As FractalTrendLineStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                            Dim filename As String = String.Format("{0} Trendline", rule.TypeOfTrendLine)

                            Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                        End Using
                    Next
#End Region
                Case 2
#Region "MarketPlusMarketMinus"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 10
                            tick = 0.05
                        Case Trade.TypeOfStock.Commodity
                            database = Common.DataBaseTable.Intraday_Commodity
                            margin = 70
                            tick = 1
                        Case Trade.TypeOfStock.Currency
                            database = Common.DataBaseTable.Intraday_Currency
                            margin = 98
                            tick = 0.0025
                        Case Trade.TypeOfStock.Futures
                            database = Common.DataBaseTable.Intraday_Futures
                            margin = 30
                            tick = 0.05
                    End Select

                    Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                            exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                            exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                            tradeStartTime:=TimeSpan.Parse("9:20:00"),
                                                            lastTradeEntryTime:=TimeSpan.Parse("14:29:59"),
                                                            eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                            tickSize:=tick,
                                                            marginMultiplier:=margin,
                                                            timeframe:=5,
                                                            heikenAshiCandle:=False,
                                                            stockType:=stockType,
                                                            databaseTable:=database,
                                                            dataSource:=sourceData,
                                                            initialCapital:=Decimal.MaxValue / 2,
                                                            usableCapital:=Decimal.MaxValue / 2,
                                                            minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                            amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Pre Market Cash Stocks.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = False
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New MarketPlusMarketMinusStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 100
                                            }

                            .NumberOfTradeableStockPerDay = Integer.MaxValue

                            .NumberOfTradesPerStockPerDay = 1

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = False
                            .StockMaxProfitPerDay = Decimal.MaxValue
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = True
                            .OverAllProfitPerDay = 15000
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim filename As String = String.Format("Market Plus Market Minus")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 3
#Region "Highest Lowest Point"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 13
                            tick = 0.05
                        Case Trade.TypeOfStock.Commodity
                            database = Common.DataBaseTable.Intraday_Commodity
                            margin = 70
                            tick = 1
                        Case Trade.TypeOfStock.Currency
                            database = Common.DataBaseTable.Intraday_Currency
                            margin = 98
                            tick = 0.0025
                        Case Trade.TypeOfStock.Futures
                            database = Common.DataBaseTable.Intraday_Futures
                            margin = 30
                            tick = 0.05
                    End Select

                    Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                            exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                            exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                            tradeStartTime:=TimeSpan.Parse("9:15:00"),
                                                            lastTradeEntryTime:=TimeSpan.Parse("14:29:59"),
                                                            eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                            tickSize:=tick,
                                                            marginMultiplier:=margin,
                                                            timeframe:=1,
                                                            heikenAshiCandle:=False,
                                                            stockType:=stockType,
                                                            databaseTable:=database,
                                                            dataSource:=sourceData,
                                                            initialCapital:=Decimal.MaxValue / 2,
                                                            usableCapital:=Decimal.MaxValue / 2,
                                                            minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                            amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Cash Stock.csv")

                            .AllowBothDirectionEntryAtSameTime = True
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New HighestLowestPointStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .ATRMultiplier = 2
                                            }

                            .NumberOfTradeableStockPerDay = 5

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = False
                            .StockMaxProfitPerDay = Decimal.MaxValue
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim filename As String = String.Format("Highest Lowest Point")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 4
#Region "Heikenashi Reversal"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 13
                            tick = 0.05
                        Case Trade.TypeOfStock.Commodity
                            database = Common.DataBaseTable.Intraday_Commodity
                            margin = 70
                            tick = 1
                        Case Trade.TypeOfStock.Currency
                            database = Common.DataBaseTable.Intraday_Currency
                            margin = 98
                            tick = 0.0025
                        Case Trade.TypeOfStock.Futures
                            database = Common.DataBaseTable.Intraday_Futures
                            margin = 30
                            tick = 0.05
                    End Select

                    Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                    tradeStartTime:=TimeSpan.Parse("9:16:00"),
                                                                    lastTradeEntryTime:=TimeSpan.Parse("14:29:59"),
                                                                    eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                    tickSize:=tick,
                                                                    marginMultiplier:=margin,
                                                                    timeframe:=1,
                                                                    heikenAshiCandle:=False,
                                                                    stockType:=stockType,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Stock.csv")

                            .AllowBothDirectionEntryAtSameTime = True
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New HeikenashiReverseSlabStrategyRule.StrategyRuleEntities With
                                                {
                                                 .SatelliteTradeTargetMultiplier = 2
                                                }

                            .NumberOfTradeableStockPerDay = 1

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = False
                            .StockMaxProfitPerDay = Decimal.MaxValue
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim filename As String = String.Format("Heikenashi Reverse Slab")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 5
#Region "EMA Scalping"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 10
                            tick = 0.05
                        Case Trade.TypeOfStock.Commodity
                            database = Common.DataBaseTable.Intraday_Commodity
                            margin = 70
                            tick = 1
                        Case Trade.TypeOfStock.Currency
                            database = Common.DataBaseTable.Intraday_Currency
                            margin = 98
                            tick = 0.0025
                        Case Trade.TypeOfStock.Futures
                            database = Common.DataBaseTable.Intraday_Futures
                            margin = 30
                            tick = 0.05
                    End Select

                    Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                    tradeStartTime:=TimeSpan.Parse("9:15:00"),
                                                                    lastTradeEntryTime:=TimeSpan.Parse("14:29:59"),
                                                                    eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                    tickSize:=tick,
                                                                    marginMultiplier:=margin,
                                                                    timeframe:=1,
                                                                    heikenAshiCandle:=False,
                                                                    stockType:=stockType,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Cash Stock.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = False
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New EMAScalpingStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 2.1,
                                                .BreakevenMovement = True
                                            }

                            .NumberOfTradeableStockPerDay = 5

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = False
                            .StockMaxProfitPerDay = Decimal.MaxValue
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim filename As String = String.Format("EMA Scalping")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 6
#Region "Supertrend Cut Reversal"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 10
                            tick = 0.05
                        Case Trade.TypeOfStock.Commodity
                            database = Common.DataBaseTable.Intraday_Commodity
                            margin = 70
                            tick = 1
                        Case Trade.TypeOfStock.Currency
                            database = Common.DataBaseTable.Intraday_Currency
                            margin = 98
                            tick = 0.0025
                        Case Trade.TypeOfStock.Futures
                            database = Common.DataBaseTable.Intraday_Futures
                            margin = 30
                            tick = 0.05
                    End Select

                    For tgtMul As Decimal = 2.1 To 2.1
                        For slAtrMul As Decimal = 2 To 3
                            For tgtAtrMul As Decimal = 1 To 2
                                For brkEvn As Integer = 0 To 1
                                    Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                                    tradeStartTime:=TimeSpan.Parse("9:17:00"),
                                                                                    lastTradeEntryTime:=TimeSpan.Parse("14:29:59"),
                                                                                    eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                                    tickSize:=tick,
                                                                                    marginMultiplier:=margin,
                                                                                    timeframe:=1,
                                                                                    heikenAshiCandle:=False,
                                                                                    stockType:=stockType,
                                                                                    databaseTable:=database,
                                                                                    dataSource:=sourceData,
                                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                                    amountToBeWithdrawn:=0)
                                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                                        With backtestStrategy
                                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "BANKNIFTY.csv")

                                            .AllowBothDirectionEntryAtSameTime = False
                                            .TrailingStoploss = False
                                            .TickBasedStrategy = False
                                            .RuleNumber = ruleNumber

                                            .RuleEntityData = New SupertrendCutReversalStrategyRule.StrategyRuleEntities With
                                                            {
                                                                .TargetMultiplier = tgtMul,
                                                                .BreakevenMovement = brkEvn,
                                                                .StoplossATRMultiplier = slAtrMul,
                                                                .TargetATRMultiplier = tgtAtrMul
                                                            }

                                            .NumberOfTradeableStockPerDay = 1

                                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                                            .ExitOnStockFixedTargetStoploss = False
                                            .StockMaxProfitPerDay = Decimal.MaxValue
                                            .StockMaxLossPerDay = Decimal.MinValue

                                            .ExitOnOverAllFixedTargetStoploss = False
                                            .OverAllProfitPerDay = Decimal.MaxValue
                                            .OverAllLossPerDay = Decimal.MinValue

                                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                                            .MovementSlab = .MTMSlab / 2
                                            .RealtimeTrailingPercentage = 50
                                        End With

                                        Dim ruleData As SupertrendCutReversalStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                                        Dim filename As String = String.Format("Supertrend Cut Reversal,TgtMul {0},SlAtrMul {1},TgtAtrMul {2},BrkEvn {3}",
                                                                               ruleData.TargetMultiplier,
                                                                               ruleData.StoplossATRMultiplier,
                                                                               ruleData.TargetATRMultiplier,
                                                                               ruleData.BreakevenMovement)

                                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                                    End Using
                                Next
                            Next
                        Next
                    Next
#End Region
                Case 7
#Region "Martingale"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 10
                            tick = 0.05
                        Case Trade.TypeOfStock.Commodity
                            database = Common.DataBaseTable.Intraday_Commodity
                            margin = 70
                            tick = 1
                        Case Trade.TypeOfStock.Currency
                            database = Common.DataBaseTable.Intraday_Currency
                            margin = 98
                            tick = 0.0025
                        Case Trade.TypeOfStock.Futures
                            database = Common.DataBaseTable.Intraday_Futures
                            margin = 30
                            tick = 0.05
                    End Select

                    Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                    tradeStartTime:=TimeSpan.Parse("9:16:00"),
                                                                    lastTradeEntryTime:=TimeSpan.Parse("14:29:59"),
                                                                    eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                    tickSize:=tick,
                                                                    marginMultiplier:=margin,
                                                                    timeframe:=1,
                                                                    heikenAshiCandle:=False,
                                                                    stockType:=stockType,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "BANKNIFTY.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = False
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New MartingaleStrategyRule.StrategyRuleEntities With
                                            {
                                                .StoplossATRMultiplier = 1 / 2,
                                                .TargetMultiplier = 4,
                                                .MinimumStoplossPoint = 25
                                            }

                            .NumberOfTradeableStockPerDay = 1

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = False
                            .StockMaxProfitPerDay = Decimal.MaxValue
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As MartingaleStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Martingale,SlAtrMul {0},TgtMul {1},MinSL {2}",
                                                               ruleData.StoplossATRMultiplier,
                                                               ruleData.TargetMultiplier,
                                                               ruleData.MinimumStoplossPoint)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 8
#Region "HL_LH Breakout"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 10
                            tick = 0.05
                        Case Trade.TypeOfStock.Commodity
                            database = Common.DataBaseTable.Intraday_Commodity
                            margin = 70
                            tick = 1
                        Case Trade.TypeOfStock.Currency
                            database = Common.DataBaseTable.Intraday_Currency
                            margin = 98
                            tick = 0.0025
                        Case Trade.TypeOfStock.Futures
                            database = Common.DataBaseTable.Intraday_Futures
                            margin = 30
                            tick = 0.05
                    End Select

                    For tf As Integer = 1 To 1
                        For signalTyp As Decimal = 0 To 1           'Signal Type:0(for Two candle), Signal Type:1(for Three Candle)
                            For tgtMul As Decimal = 4 To 4
                                For brkevnMvmnt As Integer = 0 To 0
                                    For imdtBrkout As Integer = 0 To 0
                                        Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                                        exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                                        exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                                        tradeStartTime:=TimeSpan.Parse("9:15:00"),
                                                                                        lastTradeEntryTime:=TimeSpan.Parse("14:29:59"),
                                                                                        eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                                        tickSize:=tick,
                                                                                        marginMultiplier:=margin,
                                                                                        timeframe:=tf,
                                                                                        heikenAshiCandle:=False,
                                                                                        stockType:=stockType,
                                                                                        databaseTable:=database,
                                                                                        dataSource:=sourceData,
                                                                                        initialCapital:=300000,
                                                                                        usableCapital:=200000,
                                                                                        minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                                        amountToBeWithdrawn:=0)
                                            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                                            With backtestStrategy
                                                .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Cash Stock.csv")

                                                .AllowBothDirectionEntryAtSameTime = False
                                                .TrailingStoploss = False
                                                .TickBasedStrategy = True
                                                .RuleNumber = ruleNumber

                                                .RuleEntityData = New HL_LHBreakoutStrategyRule.StrategyRuleEntities With
                                                                {
                                                                    .ATRMultiplier = 1,
                                                                    .TypeOfSignal = signalTyp,
                                                                    .TargetMultiplier = tgtMul,
                                                                    .MaxLossPerTrade = -500,
                                                                    .BreakevenMovement = brkevnMvmnt,
                                                                    .ImmediateBreakout = imdtBrkout
                                                                }

                                                .NumberOfTradeableStockPerDay = 5

                                                .NumberOfTradesPerStockPerDay = 1

                                                .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                                                .StockMaxLossPercentagePerDay = Decimal.MinValue

                                                .ExitOnStockFixedTargetStoploss = False
                                                .StockMaxProfitPerDay = Decimal.MaxValue
                                                .StockMaxLossPerDay = Decimal.MinValue

                                                .ExitOnOverAllFixedTargetStoploss = False
                                                .OverAllProfitPerDay = Decimal.MaxValue
                                                .OverAllLossPerDay = Decimal.MinValue

                                                .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                                                .MTMSlab = Math.Abs(.OverAllLossPerDay)
                                                .MovementSlab = .MTMSlab / 2
                                                .RealtimeTrailingPercentage = 50
                                            End With

                                            Dim ruleData As HL_LHBreakoutStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                                            Dim filename As String = String.Format("HLLH Output,TF {0},SgnlTyp {1},TgtMul {2},BrkevnMvmnt {3},ImdtBrkout {4}",
                                                                                   backtestStrategy.SignalTimeFrame,
                                                                                   ruleData.TypeOfSignal,
                                                                                   ruleData.TargetMultiplier,
                                                                                   ruleData.BreakevenMovement,
                                                                                   ruleData.ImmediateBreakout)

                                            Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                                        End Using
                                    Next
                                Next
                            Next
                        Next
                    Next
#End Region
                Case 9
#Region "Always in trade Martingale"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 10
                            tick = 0.05
                        Case Trade.TypeOfStock.Commodity
                            database = Common.DataBaseTable.Intraday_Commodity
                            margin = 70
                            tick = 1
                        Case Trade.TypeOfStock.Currency
                            database = Common.DataBaseTable.Intraday_Currency
                            margin = 98
                            tick = 0.0025
                        Case Trade.TypeOfStock.Futures
                            database = Common.DataBaseTable.Intraday_Futures
                            margin = 30
                            tick = 0.05
                    End Select

                    Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                    tradeStartTime:=TimeSpan.Parse("9:16:00"),
                                                                    lastTradeEntryTime:=TimeSpan.Parse("14:29:59"),
                                                                    eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                    tickSize:=tick,
                                                                    marginMultiplier:=margin,
                                                                    timeframe:=1,
                                                                    heikenAshiCandle:=False,
                                                                    stockType:=stockType,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Test.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = False
                            .RuleNumber = ruleNumber

                            .RuleEntityData = Nothing

                            .NumberOfTradeableStockPerDay = 5

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = False
                            .StockMaxProfitPerDay = Decimal.MaxValue
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim filename As String = String.Format("Always in trade Martingale")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case Else
                    Throw New NotImplementedException
            End Select

            'Delete Directory
            Dim directoryName As String = Path.Combine(My.Application.Info.DirectoryPath, String.Format("STRATEGY{0} CANDLE DATA", ruleNumber))
            'If Directory.Exists(directoryName) Then
            '    Directory.Delete(directoryName, True)
            'End If
        Catch cex As OperationCanceledException
            MsgBox(cex.Message, MsgBoxStyle.Exclamation)
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        Finally
            OnHeartbeat("Process Complete")
            SetObjectEnableDisable_ThreadSafe(btnStart, True)
            SetObjectEnableDisable_ThreadSafe(btnStop, False)
        End Try
    End Function

    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        _canceller.Cancel()
    End Sub

    Private Sub rdbDatabase_CheckedChanged(sender As Object, e As EventArgs) Handles rdbDatabase.CheckedChanged
        If rdbDatabase.Checked Then
            grpBxDBConnection.Visible = True
        End If
    End Sub

    Private Sub rdbLive_CheckedChanged(sender As Object, e As EventArgs) Handles rdbLive.CheckedChanged
        If rdbLive.Checked Then
            grpBxDBConnection.Visible = False
        End If
    End Sub
End Class

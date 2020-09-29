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
        Dim ruleNumber As Integer = GetComboBoxIndex_ThreadSafe(cmbRule)
        Try
            Dim startDate As Date = GetDateTimePickerValue_ThreadSafe(dtpckrStartDate)
            Dim endDate As Date = GetDateTimePickerValue_ThreadSafe(dtpckrEndDate)
            Dim sourceData As Strategy.SourceOfData = Strategy.SourceOfData.None
            If GetRadioButtonChecked_ThreadSafe(rdbLive) Then
                sourceData = Strategy.SourceOfData.Live
            Else
                sourceData = Strategy.SourceOfData.Database
            End If

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                                                                        optionStockType:=Trade.TypeOfStock.None,
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                                                                                    optionStockType:=Trade.TypeOfStock.None,
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
#Region "BNF Martingale"
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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

                            .RuleEntityData = New BNFMartingaleStrategyRule.StrategyRuleEntities With
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

                        Dim ruleData As BNFMartingaleStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
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
                                                                                        optionStockType:=Trade.TypeOfStock.None,
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Stock.csv")

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
                Case 10
#Region "Martingale"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            If .StockType = Trade.TypeOfStock.Cash Then
                                .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Cash Stock.csv")
                            Else
                                .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Stock.csv")
                            End If

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = False
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New MartingaleStrategyRule.StrategyRuleEntities With
                                            {
                                                .ATRMultiplier = 1,
                                                .MaxProfitPerStock = 500,
                                                .ReverseSignalExit = True
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

                        Dim ruleData As MartingaleStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("VWAP Martingale,AtrMul {0},Rvrs Sgnl {1}",
                                                               ruleData.ATRMultiplier,
                                                               ruleData.ReverseSignalExit)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 11
#Region "Anchor Satellite HK"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = False
                            .RuleNumber = ruleNumber

                            .RuleEntityData = Nothing

                            .NumberOfTradeableStockPerDay = 5

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 500
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim filename As String = String.Format("Anchor Sattelite HK Strategy Output")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 12
#Region "Small Opening Range Breakout"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
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
                                                                    timeframe:=60,
                                                                    heikenAshiCandle:=False,
                                                                    stockType:=stockType,
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Highest Volume Per Range Stock Of X Minute.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New SmallOpeningRangeBreakoutStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 2,
                                                .BreakevenMovement = False
                                            }

                            .NumberOfTradeableStockPerDay = 10

                            .NumberOfTradesPerStockPerDay = 2

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

                        Dim ruleData As SmallOpeningRangeBreakoutStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Sml Opng Rng Brkot, MxLsTrd {0}, TgtMul {1}, BrkEvnMvmnt{2}",
                                                               ruleData.MaxLossPerTrade, ruleData.TargetMultiplier, ruleData.BreakevenMovement)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 13
#Region "Loss Makeup Favourable Fractal Breakout"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New LossMakeupFavourableFractalBreakoutStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .MaxProfitPerTrade = 1000,
                                                .MinimumTargetATRMultipler = 1,
                                                .MaximumStoplossATRMultipler = 2,
                                                .MinimumStoplossATRMultipler = 1
                                            }

                            .NumberOfTradeableStockPerDay = 5

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 1000
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As LossMakeupFavourableFractalBreakoutStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Ls Mkup Frctl Brkot,MxLsTrd {0},MxPrftTrd {1},MinTgtATRMul {2},MinSlATRMul {3},MaxSlATRMul {4}",
                                                               ruleData.MaxLossPerTrade,
                                                               ruleData.MaxProfitPerTrade,
                                                               ruleData.MinimumTargetATRMultipler,
                                                               ruleData.MinimumStoplossATRMultipler,
                                                               ruleData.MaximumStoplossATRMultipler)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 14
#Region "HK Reverse Slab Martingale"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New HKReverseSlabMartingaleStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .MaxProfitPerTrade = 500
                                            }

                            .NumberOfTradeableStockPerDay = 5

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 500
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As HKReverseSlabMartingaleStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("HK Rvs Slb Martingale,MxLsTrd {0},MxPrftTrd {1}",
                                                               ruleData.MaxLossPerTrade,
                                                               ruleData.MaxProfitPerTrade)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 15
#Region "Low Price Option Buy Only Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 2
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
                                                                    optionStockType:=Trade.TypeOfStock.Futures,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Lower Price Stocks With Volume OI.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = Nothing

                            .NumberOfTradeableStockPerDay = 4

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

                        Dim filename As String = String.Format("Lower Price Option Buy Only Strategy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 16
#Region "Low Price Option OI Change Buy Only Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 2
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
                                                                    optionStockType:=Trade.TypeOfStock.Futures,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Lower Price Options With OI Change%.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = Nothing

                            .NumberOfTradeableStockPerDay = 2

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = False
                            .StockMaxProfitPerDay = Decimal.MaxValue
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = True
                            .OverAllProfitPerDay = 1000
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim filename As String = String.Format("Lower Price Option OI Change Buy Only Strategy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 17
#Region "Low Price Option Buy Only EOD Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 2
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
                                                                    optionStockType:=Trade.TypeOfStock.Futures,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Low Price Option Buy Only EOD Stocks.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = Nothing

                            .NumberOfTradeableStockPerDay = 10

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

                        Dim filename As String = String.Format("Lower Price Option Buy Only EOD Strategy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 18
#Region "Every Minute Top Gainer Losser HK Reversal Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Top Gainer Top Looser Of Every Minute.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New EveryMinuteTopGainerLosserHKReversalStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 3,
                                                .BreakevenMovement = True
                                            }

                            .NumberOfTradeableStockPerDay = 10

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

                        Dim ruleData As EveryMinuteTopGainerLosserHKReversalStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Evry Min Tp Gnr Lsr HK Rvrsl,MxLsTrd {0},TgtMul {1},Brkevn {2}",
                                                               ruleData.MaxLossPerTrade,
                                                               ruleData.TargetMultiplier,
                                                               ruleData.BreakevenMovement)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 19
#Region "Loss Makeup Rainbow Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New LossMakeupRainbowStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerStock = -500,
                                                .MaxProfitPerStock = 500
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

                        Dim filename As String = String.Format("Loss Makeup Rainbow")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 20
#Region "Anchor Satellite Loss Makeup Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

                    For atrMul As Decimal = 1 To 1.5 Step 0.5
                        For frstTrdGpEntry As Integer = 0 To 1
                            For hlfATREntry As Integer = 0 To 1
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
                                                                                optionStockType:=Trade.TypeOfStock.None,
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

                                        .RuleEntityData = New AnchorSatelliteLossMakeupStrategyRule.StrategyRuleEntities With
                                                {.ATRMultiplier = atrMul,
                                                 .FirstTradeGapEntry = frstTrdGpEntry,
                                                 .ReEntryAfterHalfATR = hlfATREntry}

                                        .NumberOfTradeableStockPerDay = 5

                                        .NumberOfTradesPerStockPerDay = Integer.MaxValue

                                        .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                                        .StockMaxLossPercentagePerDay = Decimal.MinValue

                                        .ExitOnStockFixedTargetStoploss = True
                                        .StockMaxProfitPerDay = 500
                                        .StockMaxLossPerDay = Decimal.MinValue

                                        .ExitOnOverAllFixedTargetStoploss = False
                                        .OverAllProfitPerDay = Decimal.MaxValue
                                        .OverAllLossPerDay = Decimal.MinValue

                                        .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                                        .MTMSlab = Math.Abs(.OverAllLossPerDay)
                                        .MovementSlab = .MTMSlab / 2
                                        .RealtimeTrailingPercentage = 50
                                    End With

                                    Dim ruleData As AnchorSatelliteLossMakeupStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                                    Dim filename As String = String.Format("Anchor Satellite Loss Makeu,ATRMul {0},FrstTrdGpEntry {1},HlfAtrNtry {2}",
                                                                               ruleData.ATRMultiplier,
                                                                               ruleData.FirstTradeGapEntry,
                                                                               ruleData.ReEntryAfterHalfATR)

                                    Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                                End Using
                            Next
                        Next
                    Next
#End Region
                Case 21
#Region "Anchor Satellite Loss Makeup HK Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

                    For atrMul As Decimal = 1 To 1 Step 1
                        For frstTrdMktEntry As Integer = 0 To 0
                            For hlfATREntry As Integer = 0 To 1
                                For crntATRTgt As Integer = 0 To 1
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
                                                                                    optionStockType:=Trade.TypeOfStock.None,
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

                                            .RuleEntityData = New AnchorSatelliteLossMakeupHKStrategyRule.StrategyRuleEntities With
                                                {.ATRMultiplier = atrMul,
                                                 .FirstTradeMarketEntry = frstTrdMktEntry,
                                                 .ReEntryAfterHalfATR = hlfATREntry,
                                                 .CurrentATRTarget = crntATRTgt}

                                            .NumberOfTradeableStockPerDay = 5

                                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                                            .ExitOnStockFixedTargetStoploss = True
                                            .StockMaxProfitPerDay = 500
                                            .StockMaxLossPerDay = Decimal.MinValue

                                            .ExitOnOverAllFixedTargetStoploss = False
                                            .OverAllProfitPerDay = Decimal.MaxValue
                                            .OverAllLossPerDay = Decimal.MinValue

                                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                                            .MovementSlab = .MTMSlab / 2
                                            .RealtimeTrailingPercentage = 50
                                        End With

                                        Dim ruleData As AnchorSatelliteLossMakeupHKStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                                        Dim filename As String = String.Format("Anchor Satellite Loss Makeup HK,ATRMul {0},FrstTrdMktEntry {1},HlfAtrNtry {2},CurrentATRTgt {3}",
                                                                               ruleData.ATRMultiplier,
                                                                               ruleData.FirstTradeMarketEntry,
                                                                               ruleData.ReEntryAfterHalfATR,
                                                                               ruleData.CurrentATRTarget)

                                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                                    End Using
                                Next
                            Next
                        Next
                    Next
#End Region
                Case 22
#Region "Loss Makeup Neutral Slab"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New LossMakeupNeutralSlabStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -250,
                                                .TargetMultiplier = 2
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

                        Dim ruleData As LossMakeupNeutralSlabStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Neutral Slab,MxLsTrd {0},TgtMul {1}",
                                                               ruleData.MaxLossPerTrade,
                                                               ruleData.TargetMultiplier)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 23
#Region "Neutral Slab Martingale"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New NeutralSlabMartingaleStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .MaxProfitPerTrade = 500
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

                        Dim ruleData As NeutralSlabMartingaleStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Neutral Slab,MxLsTrd {0},MxPrftTrd {1}",
                                                               ruleData.MaxLossPerTrade,
                                                               ruleData.MaxProfitPerTrade)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 24
#Region "Anchor Satellite Loss Makeup HK Futures Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 75
                            tick = 0.05
                    End Select

                    For atrMul As Decimal = 1 To 1 Step 1
                        For frstTrdMktEntry As Integer = 0 To 1
                            For hlfATREntry As Integer = 0 To 1
                                For crntATRTgt As Integer = 0 To 1
                                    For mrThnDbl As Integer = 0 To 1
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
                                                                                        optionStockType:=Trade.TypeOfStock.None,
                                                                                        databaseTable:=database,
                                                                                        dataSource:=sourceData,
                                                                                        initialCapital:=Decimal.MaxValue / 2,
                                                                                        usableCapital:=Decimal.MaxValue / 2,
                                                                                        minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                                        amountToBeWithdrawn:=0)
                                            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                                            With backtestStrategy
                                                .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "BANKNIFTY.csv")

                                                .AllowBothDirectionEntryAtSameTime = True
                                                .TrailingStoploss = False
                                                .TickBasedStrategy = True
                                                .RuleNumber = ruleNumber

                                                .RuleEntityData = New AnchorSatelliteLossMakeupHKFuturesStrategyRule.StrategyRuleEntities With
                                                    {.ATRMultiplier = atrMul,
                                                     .FirstTradeMarketEntry = frstTrdMktEntry,
                                                     .ReEntryAfterHalfATR = hlfATREntry,
                                                     .CurrentATRTarget = crntATRTgt,
                                                     .LossMakeupWithMoreThanDoubleQuantity = mrThnDbl}

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

                                            Dim ruleData As AnchorSatelliteLossMakeupHKFuturesStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                                            Dim filename As String = String.Format("AncrStltLsMkpHKFt,ATRMul {0},FrstTrdMktEntry {1},HlfAtrNtry {2},CurrentATRTgt {3},DblQty {4}",
                                                                                   ruleData.ATRMultiplier,
                                                                                   ruleData.FirstTradeMarketEntry,
                                                                                   ruleData.ReEntryAfterHalfATR,
                                                                                   ruleData.CurrentATRTarget,
                                                                                   ruleData.LossMakeupWithMoreThanDoubleQuantity)

                                            Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                                        End Using
                                    Next
                                Next
                            Next
                        Next
                    Next
#End Region
                Case 25
#Region "HK Reversal Single Trade Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 20
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New HKReversalSingleTradeStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 2
                                            }

                            .NumberOfTradeableStockPerDay = 20

                            .NumberOfTradesPerStockPerDay = 1

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 1000
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = 500
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As HKReversalSingleTradeStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("HK Rvs Sngl Trd,MxLsTrd {0},TgtMul {1}",
                                                               ruleData.MaxLossPerTrade,
                                                               ruleData.TargetMultiplier)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 26
#Region "Momentum Reversal Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 20
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New MomentumReversalStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 2,
                                                .BreakevenMovement = True
                                            }

                            .NumberOfTradeableStockPerDay = 1

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = False
                            .StockMaxProfitPerDay = Decimal.MaxValue
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = True
                            .OverAllProfitPerDay = 1000
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As MomentumReversalStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("MR,MxLsTrd {0},TgtMul {1},Brkevn {2}",
                                                               ruleData.MaxLossPerTrade, ruleData.TargetMultiplier, ruleData.BreakevenMovement)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 27
#Region "Stochastic Divergence Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 20
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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

                            .RuleEntityData = New StochasticDivergenceStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 2,
                                                .TakeAnyHighLow = False
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

                        Dim ruleData As StochasticDivergenceStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("SMIDivergence,MxLsTrd {0},TgtMul {1}",
                                                               ruleData.MaxLossPerTrade, ruleData.TargetMultiplier)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 28
#Region "Pair Anchor Satellite Loss Makeup HK Strategy"

#Region "File Setup"
                    'Dim dt As DataTable = Nothing
                    'Using csvHelper As New Utilities.DAL.CSVHelper(Path.Combine(My.Application.Info.DirectoryPath, "Stock Files", "Pair Anchor Satellite.csv"), ",", _canceller)
                    '    dt = csvHelper.GetDataTableFromCSV(1)
                    'End Using
                    'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    '    Dim dataDict As Dictionary(Of Date, List(Of Tuple(Of String, String))) = New Dictionary(Of Date, List(Of Tuple(Of String, String)))
                    '    For i = 0 To dt.Rows.Count - 1
                    '        Dim rowDate As Date = dt.Rows(i).Item(0)
                    '        Dim tradingSymbol1 As String = dt.Rows(i).Item(1)
                    '        Dim tradingSymbol2 As String = dt.Rows(i).Item(2)
                    '        Dim direction1 As String = dt.Rows(i).Item(5)
                    '        Dim direction2 As String = dt.Rows(i).Item(6)

                    '        If dataDict.ContainsKey(rowDate.Date) Then
                    '            Dim dataList As List(Of Tuple(Of String, String)) = dataDict(rowDate.Date)
                    '            dataList.Add(New Tuple(Of String, String)(tradingSymbol1, direction1))
                    '            dataList.Add(New Tuple(Of String, String)(tradingSymbol2, direction2))
                    '        Else
                    '            Dim dataList As List(Of Tuple(Of String, String)) = New List(Of Tuple(Of String, String))
                    '            dataList.Add(New Tuple(Of String, String)(tradingSymbol1, direction1))
                    '            dataList.Add(New Tuple(Of String, String)(tradingSymbol2, direction2))
                    '            dataDict.Add(rowDate.Date, dataList)
                    '        End If
                    '    Next

                    '    If dataDict IsNot Nothing AndAlso dataDict.Count > 0 Then
                    '        Dim retDt As DataTable = New DataTable
                    '        retDt.Columns.Add("Date")
                    '        retDt.Columns.Add("Trading Symbol")
                    '        retDt.Columns.Add("Direction")

                    '        For Each runningDate In dataDict
                    '            For Each runningData In runningDate.Value
                    '                Dim row As DataRow = retDt.NewRow
                    '                row("Date") = runningDate.Key.ToString("dd-MMM-yyyy")
                    '                row("Trading Symbol") = runningData.Item1
                    '                row("Direction") = runningData.Item2

                    '                retDt.Rows.Add(row)
                    '            Next
                    '        Next

                    '        Using csvHelper As New Utilities.DAL.CSVHelper(Path.Combine(My.Application.Info.DirectoryPath, "Pair Anchor Satellite.csv"), ",", _canceller)
                    '            csvHelper.GetCSVFromDataTable(retDt)
                    '        End Using
                    '    End If
                    'End If
#End Region

                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

                    For atrMul As Decimal = 1 To 1 Step 1
                        For frstTrdMktEntry As Integer = 0 To 0
                            For hlfATREntry As Integer = 1 To 1
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
                                                                                optionStockType:=Trade.TypeOfStock.None,
                                                                                databaseTable:=database,
                                                                                dataSource:=sourceData,
                                                                                initialCapital:=Decimal.MaxValue / 2,
                                                                                usableCapital:=Decimal.MaxValue / 2,
                                                                                minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                                amountToBeWithdrawn:=0)
                                    AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                                    With backtestStrategy
                                        .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Pair Anchor Satellite.csv")

                                        .AllowBothDirectionEntryAtSameTime = True
                                        .TrailingStoploss = False
                                        .TickBasedStrategy = True
                                        .RuleNumber = ruleNumber

                                        .RuleEntityData = New PairAnchorHKStrategyRule.StrategyRuleEntities With
                                            {.FirstTradeMarketEntry = frstTrdMktEntry}

                                        .NumberOfTradeableStockPerDay = 2

                                        .NumberOfTradesPerStockPerDay = Integer.MaxValue

                                        .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                                        .StockMaxLossPercentagePerDay = Decimal.MinValue

                                        .ExitOnStockFixedTargetStoploss = False
                                        .StockMaxProfitPerDay = Decimal.MaxValue
                                        .StockMaxLossPerDay = Decimal.MinValue

                                        .ExitOnOverAllFixedTargetStoploss = True
                                        .OverAllProfitPerDay = 1000
                                        .OverAllLossPerDay = Decimal.MinValue

                                        .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                                        .MTMSlab = Math.Abs(.OverAllLossPerDay)
                                        .MovementSlab = .MTMSlab / 2
                                        .RealtimeTrailingPercentage = 50
                                    End With

                                    Dim ruleData As PairAnchorHKStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                                    Dim filename As String = String.Format("Pair Anchor HK,FrstTrdMktEntry {0}",
                                                                           ruleData.FirstTradeMarketEntry)

                                    Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                                End Using
                            Next
                        Next
                    Next
#End Region
                Case 29
#Region "Two Third Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 20
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Low ATR Candle Quick Entry Stocks.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = False
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New TwoThirdStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -100,
                                                .TargetMultiplier = 5,
                                                .MaxATRMultiplier = 1,
                                                .MinATRMultiplier = 0.5
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

                        Dim ruleData As TwoThirdStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("TwoThirdStrategy,MxLsTrd {0},TgtMul {1}",
                                                               ruleData.MaxLossPerTrade, ruleData.TargetMultiplier)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 30
#Region "Small Range Breakout"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Low ATR Candle Quick Entry Stocks.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New SmallRangeBreakoutStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 2,
                                                .BreakevenMovement = False,
                                                .ReverseSignalEntry = False
                                            }

                            .NumberOfTradeableStockPerDay = 20

                            .NumberOfTradesPerStockPerDay = 2

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

                        Dim ruleData As SmallRangeBreakoutStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Sml Rng Brkot, MxLsTrd {0}, TgtMul {1}, BrkEvnMvmnt {2}, Rvrs {3}",
                                                               ruleData.MaxLossPerTrade, ruleData.TargetMultiplier, ruleData.BreakevenMovement, ruleData.ReverseSignalEntry)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 31
#Region "Highest Lowest Point Anchor Satellite Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

                    Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                    tradeStartTime:=TimeSpan.Parse("9:15:00"),
                                                                    lastTradeEntryTime:=TimeSpan.Parse("14:44:59"),
                                                                    eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                    tickSize:=tick,
                                                                    marginMultiplier:=margin,
                                                                    timeframe:=1,
                                                                    heikenAshiCandle:=False,
                                                                    stockType:=stockType,
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "EOD Low Range Stocks.csv")

                            .AllowBothDirectionEntryAtSameTime = True
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New HighestLowestPointAnchorSatelliteStrategyRule.StrategyRuleEntities With
                                    {.ATRMultiplier = 1}

                            .NumberOfTradeableStockPerDay = 5

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 500
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = True
                            .OverAllProfitPerDay = 2000
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As HighestLowestPointAnchorSatelliteStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Highest Lowest Point Anchor Satellite Strategy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 32
#Region "EMA SMA Crossover Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

                    Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                    tradeStartTime:=TimeSpan.Parse("9:17:00"),
                                                                    lastTradeEntryTime:=TimeSpan.Parse("14:44:59"),
                                                                    eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                    tickSize:=tick,
                                                                    marginMultiplier:=margin,
                                                                    timeframe:=1,
                                                                    heikenAshiCandle:=False,
                                                                    stockType:=stockType,
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "EOD Low Range Stocks.csv")

                            .AllowBothDirectionEntryAtSameTime = True
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New EmaSmaCrossoverStrategyRule.StrategyRuleEntities With
                                    {.MaxLossPerTrade = -500, .TargetMultiplier = 2}

                            .NumberOfTradeableStockPerDay = 20

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 1000
                            .StockMaxLossPerDay = -1000

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As EmaSmaCrossoverStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("EMA SMA Crossover Strategy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 33
#Region "Bollinger Close Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 20
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
                            margin = 45
                            tick = 0.05
                    End Select

                    For tf As Integer = 1 To 1
                        For atrMul As Decimal = 1 To 1
                            For sd As Decimal = 3 To 3
                                For prftExt As Integer = 1 To 1
                                    Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                                    tradeStartTime:=TimeSpan.Parse("9:16:00"),
                                                                                    lastTradeEntryTime:=TimeSpan.Parse("14:44:59"),
                                                                                    eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                                    tickSize:=tick,
                                                                                    marginMultiplier:=margin,
                                                                                    timeframe:=tf,
                                                                                    heikenAshiCandle:=False,
                                                                                    stockType:=stockType,
                                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                                            .TickBasedStrategy = True
                                            .RuleNumber = ruleNumber

                                            .RuleEntityData = New BollingerCloseStrategyRule.StrategyRuleEntities With
                                                {.MaxLossPerTrade = -500,
                                                 .ATRMultiplier = atrMul,
                                                 .BollingerPeriod = 20,
                                                 .StandardDeviation = sd,
                                                 .ExitAtProfit = prftExt}

                                            .NumberOfTradeableStockPerDay = 5

                                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                                            .ExitOnStockFixedTargetStoploss = False
                                            .StockMaxProfitPerDay = Decimal.MaxValue
                                            .StockMaxLossPerDay = Decimal.MinValue

                                            .ExitOnOverAllFixedTargetStoploss = True
                                            .OverAllProfitPerDay = 4000
                                            .OverAllLossPerDay = Decimal.MinValue

                                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                                            .MovementSlab = .MTMSlab / 2
                                            .RealtimeTrailingPercentage = 50
                                        End With

                                        Dim ruleData As BollingerCloseStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                                        Dim filename As String = String.Format("BlngrClsOtpt,TF {0},AtrMul {1},SD {2},PrftExt {3}",
                                                                               backtestStrategy.SignalTimeFrame,
                                                                               ruleData.ATRMultiplier,
                                                                               ruleData.StandardDeviation,
                                                                               ruleData.ExitAtProfit)

                                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                                    End Using
                                Next
                            Next
                        Next
                    Next
#End Region
                Case 34
#Region "Outside VWAP Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "EOD Low Range Stocks.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New OutsideVWAPStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500
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

                        Dim ruleData As OutsideVWAPStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Outside VWAP Strategy Output")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 35
#Region "Higher Timeframe Direction Martingale Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New HigherTimeframeDirectionMartingaleStrategyRule.StrategyRuleEntities With
                                            {
                                                .HigherTimeframe = 5,
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 1
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

                        Dim ruleData As HigherTimeframeDirectionMartingaleStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("HTDirMrtnglStrtgy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 36
#Region "Squeeze Breakout Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New SqueezeBreakoutStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500
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

                        Dim ruleData As SqueezeBreakoutStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("SqzBrkotStrtgy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 37
#Region "Multi Trade Loss Makeup Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

                    For tgtMul As Decimal = 4 To 4
                        For brkEvnMvmnt As Integer = 1 To 1
                            For tgtMd As Integer = 1 To 1
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
                                                                                optionStockType:=Trade.TypeOfStock.None,
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
                                        .TickBasedStrategy = True
                                        .RuleNumber = ruleNumber

                                        .RuleEntityData = New MultiTradeLossMakeupStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -250,
                                                .TargetMultiplier = tgtMul,
                                                .BreakevenMovement = brkEvnMvmnt,
                                                .TargetMode = tgtMd,
                                                .NumberOfLossTrade = 3,
                                                .MultipleTradeInASignal = False
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

                                    Dim ruleData As MultiTradeLossMakeupStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                                    Dim filename As String = String.Format("Multi Trade Loss Makeup Strategy,TgtMul {0},BrkEvn {1},TgtMode {2}",
                                                                           ruleData.TargetMultiplier, ruleData.BreakevenMovement, ruleData.TargetMode.ToString)

                                    Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                                End Using
                            Next
                        Next
                    Next
#End Region
                Case 38
#Region "Momentum Reversal Modified Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 20
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New MomentumReversalModifiedStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 2,
                                                .BreakevenMovement = True
                                            }

                            .NumberOfTradeableStockPerDay = 5

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 1000
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As MomentumReversalModifiedStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("MR,MxLsTrd {0},TgtMul {1},Brkevn {2}",
                                                               ruleData.MaxLossPerTrade, ruleData.TargetMultiplier, ruleData.BreakevenMovement)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 39
#Region "HK Reverse Exit Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 20
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New HKReverseExitStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500
                                            }

                            .NumberOfTradeableStockPerDay = 20

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

                        Dim ruleData As HKReverseExitStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("HK Reverse Exit Strategy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 40
#Region "Both Direction Multi Trades HK Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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

                            .RuleEntityData = Nothing

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

                        Dim filename As String = String.Format("Both Direction Multi Trade HK Strategy Output")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 41
#Region "Market Entry Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 20
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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

                            .RuleEntityData = New MarketEntryStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .ATRMultiplier = 1.5,
                                                .TargetMultiplier = 3
                                            }

                            .NumberOfTradeableStockPerDay = 5

                            .NumberOfTradesPerStockPerDay = 2

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

                        Dim ruleData As MarketEntryStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Market Entry Strategy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 42
#Region "HK Reversal Loss Makeup Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 20
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New HKReversalLossMakeupStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 1
                                            }

                            .NumberOfTradeableStockPerDay = 5

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 500
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As HKReversalLossMakeupStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("HK Reversal Loss Makeup Strategy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 43
#Region "Supertrend Cut Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 20
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Cash Stock Modified.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New SupertrendCutStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -250,
                                                .TargetMultiplier = 3,
                                                .StoplossPercentage = 0.3
                                            }

                            .NumberOfTradeableStockPerDay = Integer.MaxValue

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = False
                            .StockMaxProfitPerDay = Decimal.MaxValue
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = True
                            .OverAllProfitPerDay = 500
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As SupertrendCutStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Supertrend Cut Strategy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 44
#Region "Both Direction Multi Trades Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Dummy Stock Data.csv")

                            .AllowBothDirectionEntryAtSameTime = True
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = Nothing

                            .NumberOfTradeableStockPerDay = Integer.MaxValue

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 500
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim filename As String = String.Format("Both Direction Multi Trade Strategy Output")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 45
#Region "Day High Low Swing Trendline Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = Nothing

                            .NumberOfTradeableStockPerDay = 5

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 500
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = True
                            .OverAllProfitPerDay = 500
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim filename As String = String.Format("Day High Low Swing Trendline Strategy Output")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 46
#Region "Fibonacci Backtest Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "NIFTY.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New FibonacciBacktestStrategyRule.StrategyRuleEntities With {.Multiplier = 1}

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

                        Dim filename As String = String.Format("Fibonacci Strategy Output")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 47
#Region "Swing At Day High Low Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

                    For tgtMul As Decimal = 2 To 3
                        For brkevn As Integer = 0 To 1
                            For ovrlPrft As Decimal = 1000 To 3000 Step 1000
                                For ovrlLs As Decimal = -1000 To -3000 Step -1000
                                    If Math.Abs(ovrlLs) > ovrlPrft Then Continue For

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                                        With backtestStrategy
                                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Cash Stock Modified Price Range.csv")

                                            .AllowBothDirectionEntryAtSameTime = False
                                            .TrailingStoploss = False
                                            .TickBasedStrategy = False
                                            .RuleNumber = ruleNumber

                                            .RuleEntityData = New SwingAtDayHLStrategyRule.StrategyRuleEntities With
                                                              {.ATRMultiplier = 1.5,
                                                               .MaxLossPerTrade = -500,
                                                               .TargetMultiplier = tgtMul,
                                                               .BreakevenMovement = brkevn,
                                                               .BreakevenTargetMultiplier = tgtMul / 2,
                                                               .NumberOfTradeOnEachDirection = 2}

                                            .NumberOfTradeableStockPerDay = Integer.MaxValue

                                            .NumberOfTradesPerStockPerDay = 4

                                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                                            .ExitOnStockFixedTargetStoploss = False
                                            .StockMaxProfitPerDay = Decimal.MaxValue
                                            .StockMaxLossPerDay = Decimal.MinValue

                                            .ExitOnOverAllFixedTargetStoploss = True
                                            .OverAllProfitPerDay = ovrlPrft
                                            .OverAllLossPerDay = ovrlLs

                                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                                            .MovementSlab = .MTMSlab / 2
                                            .RealtimeTrailingPercentage = 50
                                        End With

                                        Dim ruleData As SwingAtDayHLStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                                        Dim filename As String = String.Format("SwngDyHL,TgtMul {0},Brkevn {1},OvrlPrft {2},OvrlLs {3}",
                                                                               ruleData.TargetMultiplier,
                                                                               ruleData.BreakevenMovement,
                                                                               backtestStrategy.OverAllProfitPerDay,
                                                                               backtestStrategy.OverAllLossPerDay)

                                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                                    End Using
                                Next
                            Next
                        Next
                    Next
#End Region
                Case 48
#Region "Fibonacci Opening Range Breakout Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

                    Dim trgtLvls As List(Of Decimal) = New List(Of Decimal) From {38.2, 50, 61.8}
                    For Each trgt In trgtLvls
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                            With backtestStrategy
                                .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Cash Stock Modified Price Range.csv")

                                .AllowBothDirectionEntryAtSameTime = False
                                .TrailingStoploss = False
                                .TickBasedStrategy = True
                                .RuleNumber = ruleNumber

                                .RuleEntityData = New FibonacciOpeningRangeBreakoutStrategyRule.StrategyRuleEntities With
                                              {.MaxLossPerTrade = -1000,
                                               .StoplossLevel = 50,
                                               .TargetLevel = trgt}

                                .NumberOfTradeableStockPerDay = 10

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

                            Dim ruleData As FibonacciOpeningRangeBreakoutStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                            Dim filename As String = String.Format("Fibonacci Opening Range Breakout,TrgtLvl {0}", Math.Floor(ruleData.TargetLevel))

                            Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                        End Using
                    Next
#End Region
                Case 49
#Region "Previous Day HK Trend Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 20
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
                            margin = 50
                            tick = 0.05
                    End Select

                    Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                    tradeStartTime:=TimeSpan.Parse("9:15:00"),
                                                                    lastTradeEntryTime:=TimeSpan.Parse("14:44:59"),
                                                                    eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                    tickSize:=tick,
                                                                    marginMultiplier:=margin,
                                                                    timeframe:=1,
                                                                    heikenAshiCandle:=False,
                                                                    stockType:=stockType,
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Cash Stock Modified Price Range.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New PreviousDayHKTrendStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 3,
                                                .BreakevenMovement = True
                                            }

                            .NumberOfTradeableStockPerDay = 10

                            .NumberOfTradesPerStockPerDay = 3

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = False
                            .StockMaxProfitPerDay = Decimal.MaxValue
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = 500
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As PreviousDayHKTrendStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Prvs HK Trnd Strtgy,MxLsTrd {0},TgtMul {1}",
                                                               ruleData.MaxLossPerTrade,
                                                               ruleData.TargetMultiplier)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 50
#Region "Previous Day HK Trend Bollinger Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 20
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Previous Day Strong HK Stocks.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New PreviousDayHKTrendBollingerStrategyRule.StrategyRuleEntities With
                                            {
                                                .TargetPrice = 500
                                            }

                            .NumberOfTradeableStockPerDay = 10

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
                            .MTMSlab = 500
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim filename As String = String.Format("Prvs HK Trnd Blngr Strtgy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 51
#Region "EMA Attraction Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Cash Stock Modified Price Range.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New EMAAttractionStrategyRule.StrategyRuleEntities With
                                            {
                                                .EMAPeriod = 13,
                                                .ATRMultipler = 1,
                                                .MaxLossPerTrade = -500,
                                                .TakeDoubleQuantity = False
                                            }

                            .NumberOfTradeableStockPerDay = 5

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 500
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As EMAAttractionStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("EMA Attraction Strategy,EMA {0},ATRMul {1},TkDblQty {2}",
                                                               ruleData.EMAPeriod, ruleData.ATRMultipler, ruleData.TakeDoubleQuantity)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 52
#Region "Buy Below Fractal Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 1
                            tick = 0.05
                    End Select

                    Dim sortList As List(Of String) = New List(Of String) From {"Volume", "ATR"}
                    Dim filterList As List(Of String) = New List(Of String) From {"No", "Volume"}
                    Dim priceFilterList As List(Of String) = New List(Of String) From {"PreviousClose", "CurrentOpen"}
                    Dim stockTypeList As List(Of String) = New List(Of String) From {"CEPE", "Top2"}
                    For Each runningSort In sortList
                        For Each runningFilter In filterList
                            For Each runningPriceFilter In priceFilterList
                                For Each runningStockType In stockTypeList
                                    For targetAdjustment As Integer = 0 To 1
                                        Dim potentialStockFilename As String = String.Format("{0} Sort {1} Filter {2} {3}", runningSort, runningFilter, runningPriceFilter, runningStockType)
                                        Dim filepath As String = Path.Combine(My.Application.Info.DirectoryPath, String.Format("{0}.csv", potentialStockFilename))
                                        If File.Exists(filepath) Then
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
                                                                                            optionStockType:=Trade.TypeOfStock.Futures,
                                                                                            databaseTable:=database,
                                                                                            dataSource:=sourceData,
                                                                                            initialCapital:=Decimal.MaxValue / 2,
                                                                                            usableCapital:=Decimal.MaxValue / 2,
                                                                                            minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                                            amountToBeWithdrawn:=0)
                                                AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                                                With backtestStrategy
                                                    .StockFileName = filepath

                                                    .AllowBothDirectionEntryAtSameTime = False
                                                    .TrailingStoploss = False
                                                    .TickBasedStrategy = True
                                                    .RuleNumber = ruleNumber

                                                    .RuleEntityData = New BuyBelowFractalStrategyRule.StrategyRuleEntities With {.AdjustTarget = False}

                                                    .NumberOfTradeableStockPerDay = 2

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

                                                Dim ruleData As BuyBelowFractalStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                                                Dim filename As String = String.Format("{0},AdjustTarget {1}", potentialStockFilename, ruleData.AdjustTarget)

                                                Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                                            End Using
                                        End If
                                    Next
                                Next
                            Next
                        Next
                    Next
#End Region
                Case 53
#Region "HK Reversal Martingale Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Previous Day Strong HK Stocks.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New HKReversalMartingaleStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .MaxProfitPerTrade = 500
                                            }

                            .NumberOfTradeableStockPerDay = 4

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 500
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As HKReversalMartingaleStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("HK Rvs Martingale,MxLsTrd {0},MxPrftTrd {1}",
                                                               ruleData.MaxLossPerTrade,
                                                               ruleData.MaxProfitPerTrade)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 54
#Region "HK Reversal Adaptive Martingale Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Pre Market Stock 16_08_20 to 20_08_20.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New HKReversalAdaptiveMartingaleStrategyRule.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .MaxProfitPerTrade = 500
                                            }

                            .NumberOfTradeableStockPerDay = 4

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

                        Dim ruleData As HKReversalAdaptiveMartingaleStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("HK Rvs Adptv Martingale,MxLsTrd {0},MxPrftTrd {1}",
                                                               ruleData.MaxLossPerTrade,
                                                               ruleData.MaxProfitPerTrade)

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 55
#Region "Bolliger Touch Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

                    For overallLoss As Decimal = 5000 To 20000 Step 5000
                        For overallProfitMul As Decimal = 2 To 4 Step 0.5
                            If overallLoss * overallProfitMul > 40000 Then Continue For

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
                                                                            optionStockType:=Trade.TypeOfStock.None,
                                                                            databaseTable:=database,
                                                                            dataSource:=sourceData,
                                                                            initialCapital:=Decimal.MaxValue / 2,
                                                                            usableCapital:=Decimal.MaxValue / 2,
                                                                            minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                            amountToBeWithdrawn:=0)
                                AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                                With backtestStrategy
                                    .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based All Cash Stock Modified Price Range.csv")

                                    .AllowBothDirectionEntryAtSameTime = False
                                    .TrailingStoploss = False
                                    .TickBasedStrategy = True
                                    .RuleNumber = ruleNumber

                                    .RuleEntityData = New BollingerTouchStrategyRule.StrategyRuleEntities With
                                                    {
                                                        .MaxLossPerTrade = -500,
                                                        .TargetMultiplier = 4
                                                    }

                                    .NumberOfTradeableStockPerDay = 20

                                    .NumberOfTradesPerStockPerDay = Integer.MaxValue

                                    .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                                    .StockMaxLossPercentagePerDay = Decimal.MinValue

                                    .ExitOnStockFixedTargetStoploss = True
                                    .StockMaxProfitPerDay = 2000
                                    .StockMaxLossPerDay = -1000

                                    .ExitOnOverAllFixedTargetStoploss = True
                                    .OverAllProfitPerDay = overallLoss * overallProfitMul
                                    .OverAllLossPerDay = overallLoss * -1

                                    .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                                    .MTMSlab = Math.Abs(.OverAllLossPerDay)
                                    .MovementSlab = .MTMSlab / 2
                                    .RealtimeTrailingPercentage = 50
                                End With

                                Dim ruleData As BollingerTouchStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                                Dim filename As String = String.Format("Blngr Tch Strtgy,StkPrft {0},OvrlPrft {1},OvrlLs {2}",
                                                                       backtestStrategy.StockMaxProfitPerDay,
                                                                       backtestStrategy.OverAllProfitPerDay,
                                                                       backtestStrategy.OverAllLossPerDay)

                                Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                            End Using
                        Next
                    Next
#End Region
                Case 56
#Region "At The Money Option Buy Only Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 2
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
                                                                    optionStockType:=Trade.TypeOfStock.Futures,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "At The Money Option Buy Hedging Stocklist.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = Nothing

                            .NumberOfTradeableStockPerDay = 2

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

                        Dim filename As String = String.Format("At The Money Option Hedging Buy Only Strategy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 57
#Region "Previous Day HK Trend Swing Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
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
                                                                    optionStockType:=Trade.TypeOfStock.None,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                    amountToBeWithdrawn:=0)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Pre Market Stock Value sort.csv")

                            .AllowBothDirectionEntryAtSameTime = False
                            .TrailingStoploss = False
                            .TickBasedStrategy = False
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New PreviousDayHKTrendSwingStrategy.StrategyRuleEntities With
                                            {
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 2,
                                                .ATRMultiplier = 1 / 4
                                            }

                            .NumberOfTradeableStockPerDay = 20

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

                        Dim ruleData As PreviousDayHKTrendSwingStrategy.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("Prvs Dy Hk Trnd Strgy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case 58
#Region "Pre Market Options Direction Based Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 1
                            tick = 0.05
                    End Select

                    For drctn As Integer = 1 To 2
                        Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                        exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                        exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                        tradeStartTime:=TimeSpan.Parse("9:16:00"),
                                                                        lastTradeEntryTime:=TimeSpan.Parse("09:16:59"),
                                                                        eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                        tickSize:=tick,
                                                                        marginMultiplier:=margin,
                                                                        timeframe:=1,
                                                                        heikenAshiCandle:=False,
                                                                        stockType:=stockType,
                                                                        optionStockType:=Trade.TypeOfStock.Futures,
                                                                        databaseTable:=database,
                                                                        dataSource:=sourceData,
                                                                        initialCapital:=Decimal.MaxValue / 2,
                                                                        usableCapital:=Decimal.MaxValue / 2,
                                                                        minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                        amountToBeWithdrawn:=0)
                            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                            With backtestStrategy
                                .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Pre Market Top 3 Options List.csv")

                                .AllowBothDirectionEntryAtSameTime = False
                                .TrailingStoploss = False
                                .TickBasedStrategy = True
                                .RuleNumber = ruleNumber

                                .RuleEntityData = New PreMarketOptionDirectionBasedStrategy.StrategyRuleEntities With
                                {.Direction = drctn}

                                .NumberOfTradeableStockPerDay = Integer.MaxValue

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

                            Dim ruleData As PreMarketOptionDirectionBasedStrategy.StrategyRuleEntities = backtestStrategy.RuleEntityData
                            Dim filename As String = String.Format("Pre Market Options Direction Based Strategy,Directon {0}", ruleData.Direction.ToString)

                            Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                        End Using
                    Next
#End Region
                Case 59
#Region "Multi Timeframe Moving Average Strategy"
                    Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
                    Dim database As Common.DataBaseTable = Common.DataBaseTable.None
                    Dim margin As Decimal = 0
                    Dim tick As Decimal = 0
                    Select Case stockType
                        Case Trade.TypeOfStock.Cash
                            database = Common.DataBaseTable.Intraday_Cash
                            margin = 15
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
                            margin = 50
                            tick = 0.05
                    End Select

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
                                                                    optionStockType:=Trade.TypeOfStock.None,
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
                            .TickBasedStrategy = True
                            .RuleNumber = ruleNumber

                            .RuleEntityData = New MultiTimeframeMAStrategy.StrategyRuleEntities With
                                            {
                                                .HigherTimeframe = 15,
                                                .MaxLossPerTrade = -500,
                                                .TargetMultiplier = 2,
                                                .TargetInINR = True
                                            }

                            .NumberOfTradeableStockPerDay = 10

                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                            .StockMaxLossPercentagePerDay = Decimal.MinValue

                            .ExitOnStockFixedTargetStoploss = True
                            .StockMaxProfitPerDay = 1000
                            .StockMaxLossPerDay = Decimal.MinValue

                            .ExitOnOverAllFixedTargetStoploss = False
                            .OverAllProfitPerDay = Decimal.MaxValue
                            .OverAllLossPerDay = Decimal.MinValue

                            .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
                            .MTMSlab = Math.Abs(.OverAllLossPerDay)
                            .MovementSlab = .MTMSlab / 2
                            .RealtimeTrailingPercentage = 50
                        End With

                        Dim ruleData As MultiTimeframeMAStrategy.StrategyRuleEntities = backtestStrategy.RuleEntityData
                        Dim filename As String = String.Format("MultiTimeframeMAStrategy")

                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
#End Region
                Case Else
                    Throw New NotImplementedException
            End Select

            'Delete Directory
            Dim directoryName As String = Path.Combine(My.Application.Info.DirectoryPath, String.Format("STRATEGY{0} CANDLE DATA", ruleNumber))
            If Directory.Exists(directoryName) Then
                Directory.Delete(directoryName, True)
            End If
        Catch cex As OperationCanceledException
            ''Delete Directory
            'Dim directoryName As String = Path.Combine(My.Application.Info.DirectoryPath, String.Format("STRATEGY{0} CANDLE DATA", ruleNumber))
            'If Directory.Exists(directoryName) Then
            '    Directory.Delete(directoryName, True)
            'End If

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

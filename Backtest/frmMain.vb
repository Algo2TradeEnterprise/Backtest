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
        cmbRule.SelectedIndex = My.Settings.Rule
        rdbDatabase.Checked = My.Settings.Database
        rdbLive.Checked = My.Settings.Live
        rdbMIS.Checked = My.Settings.MIS
        rdbCNCTick.Checked = My.Settings.CNCTick
        rdbCNCCandle.Checked = My.Settings.CNCCandle
        rdbCNCEOD.Checked = My.Settings.CNCEOD
        If My.Settings.StartDate <> Date.MinValue Then
            dtpckrStartDate.Value = My.Settings.StartDate
        End If
        If My.Settings.EndDate <> Date.MinValue Then
            dtpckrEndDate.Value = My.Settings.EndDate
        End If
        SetObjectEnableDisable_ThreadSafe(btnStop, False)
    End Sub

    Private Async Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        _canceller = New CancellationTokenSource
        SetObjectEnableDisable_ThreadSafe(btnStart, False)
        SetObjectEnableDisable_ThreadSafe(btnStop, True)
        My.Settings.Rule = cmbRule.SelectedIndex
        My.Settings.Database = rdbDatabase.Checked
        My.Settings.Live = rdbLive.Checked
        My.Settings.MIS = rdbMIS.Checked
        My.Settings.CNCTick = rdbCNCTick.Checked
        My.Settings.CNCEOD = rdbCNCEOD.Checked
        My.Settings.CNCCandle = rdbCNCCandle.Checked
        My.Settings.StartDate = dtpckrStartDate.Value
        My.Settings.EndDate = dtpckrEndDate.Value
        My.Settings.Save()

        If rdbMIS.Checked Then
            Await Task.Run(AddressOf ViewDataMISAsync).ConfigureAwait(False)
        ElseIf rdbCNCTick.Checked Then
            Await Task.Run(AddressOf ViewDataCNCAsync).ConfigureAwait(False)
        ElseIf rdbCNCEOD.Checked Then
            Await Task.Run(AddressOf ViewDataCNCEODAsync).ConfigureAwait(False)
        ElseIf rdbCNCCandle.Checked Then
            Await Task.Run(AddressOf ViewDataCNCCandleAsync).ConfigureAwait(False)
        End If
    End Sub

#Region "MIS"
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
                    margin = 30
                    tick = 0.05
            End Select

            Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                              exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                              exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                              tradeStartTime:=TimeSpan.Parse("9:30:00"),
                                                              lastTradeEntryTime:=TimeSpan.Parse("14:45:00"),
                                                              eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                              tickSize:=tick,
                                                              marginMultiplier:=margin,
                                                              timeframe:=15,
                                                              heikenAshiCandle:=False,
                                                              stockType:=stockType,
                                                              databaseTable:=database,
                                                              dataSource:=sourceData,
                                                              initialCapital:=Decimal.MaxValue / 2,
                                                              usableCapital:=Decimal.MinValue / 2,
                                                              minimumEarnedCapitalToWithdraw:=Decimal.MinValue / 2,
                                                              amountToBeWithdrawn:=10000)
                AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                With backtestStrategy
                    .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "NSE 100 Stock List.csv")

                    .AllowBothDirectionEntryAtSameTime = False
                    .TrailingStoploss = False
                    .TickBasedStrategy = True
                    .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)

                    .RuleEntityData = New SachinPatelStrategyRule.StrategyRuleEntities With
                                {.InvestmentPerStock = 10000,
                                 .MaxAllowedLossPercentagePerStock = 3,
                                 .MaxStoplossPercentagePerTrade = 0.3,
                                 .MaxTargetPercentagePerTrade = 0.3
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

                    .TypeOfMTMTrailing = False
                    .MTMSlab = Math.Abs(.OverAllLossPerDay)
                    .MovementSlab = .MTMSlab / 2
                    .RealtimeTrailingPercentage = 50
                End With

                Dim filename As String = String.Format("Backtest Output")

                Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
            End Using
        Catch ex As Exception
            MsgBox(ex.StackTrace, MsgBoxStyle.Critical)
        Finally
            OnHeartbeat("Process Complete")
            SetObjectEnableDisable_ThreadSafe(btnStart, True)
            SetObjectEnableDisable_ThreadSafe(btnStop, False)
        End Try
    End Function
#End Region

#Region "CNC Tick"
    Private Async Function ViewDataCNCAsync() As Task
        Try
            Dim startDate As Date = GetDateTimePickerValue_ThreadSafe(dtpckrStartDate)
            Dim endDate As Date = GetDateTimePickerValue_ThreadSafe(dtpckrEndDate)
            Dim sourceData As Strategy.SourceOfData = Strategy.SourceOfData.None
            If GetRadioButtonChecked_ThreadSafe(rdbLive) Then
                sourceData = Strategy.SourceOfData.Live
            Else
                sourceData = Strategy.SourceOfData.Database
            End If
            Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
            Dim database As Common.DataBaseTable = Common.DataBaseTable.None
            Dim margin As Decimal = 0
            Dim tick As Decimal = 0
            Select Case stockType
                Case Trade.TypeOfStock.Cash
                    database = Common.DataBaseTable.Intraday_Cash
                    margin = 1
                    tick = 0.05
                Case Trade.TypeOfStock.Commodity
                    database = Common.DataBaseTable.Intraday_Commodity
                    margin = 1
                    tick = 1
                Case Trade.TypeOfStock.Currency
                    database = Common.DataBaseTable.Intraday_Currency
                    margin = 1
                    tick = 0.0025
                Case Trade.TypeOfStock.Futures
                    database = Common.DataBaseTable.Intraday_Futures
                    margin = 1
                    tick = 0.05
            End Select

            Using backtestStrategy As New CNCGenericStrategy(canceller:=_canceller,
                                                            exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                            exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                            tradeStartTime:=TimeSpan.Parse("09:15:00"),
                                                            lastTradeEntryTime:=TimeSpan.Parse("15:29:59"),
                                                            eodExitTime:=TimeSpan.Parse("15:29:59"),
                                                            tickSize:=tick,
                                                            marginMultiplier:=margin,
                                                            timeframe:=180,
                                                            heikenAshiCandle:=False,
                                                            stockType:=stockType,
                                                            databaseTable:=database,
                                                            dataSource:=sourceData,
                                                            initialCapital:=Decimal.MaxValue / 2,
                                                            usableCapital:=Decimal.MaxValue / 2,
                                                            minimumEarnedCapitalToWithdraw:=Decimal.MaxValue / 2,
                                                            amountToBeWithdrawn:=5000)
                AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                With backtestStrategy
                    '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Vijay CNC Instrument Details.csv")
                    .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Investment Stock List.csv")

                    .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)

                    Select Case .RuleNumber
                        Case 10
                            .RuleEntityData = New VijayCNCStrategyRule.StrategyRuleEntities With {.RefreshQuantityAtDayStart = False}
                        Case 18
                            .RuleEntityData = New InvestmentCNCStrategyRule.StrategyRuleEntities With {.QuantityType = InvestmentCNCStrategyRule.TypeOfQuantity.Linear}
                        Case 26
                            .RuleEntityData = New HKPositionalHourlyStrategyRule1.StrategyRuleEntities With
                                        {.QuantityType = HKPositionalHourlyStrategyRule1.TypeOfQuantity.Linear,
                                         .QuntityForLinear = 2,
                                         .TargetMultiplier = 0.5,
                                         .TypeOfExit = HKPositionalHourlyStrategyRule1.ExitType.CompoundingToMonthlyATR}
                    End Select

                    .NumberOfTradeableStockPerDay = 1

                    .NumberOfTradesPerDay = Integer.MaxValue
                    .NumberOfTradesPerStockPerDay = Integer.MaxValue

                    .TickBasedStrategy = True
                End With
                Dim filename As String = String.Format("CNC Output Capital {3} {0}_{1}_{2}", Now.Hour, Now.Minute, Now.Second,
                                                   If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital))
                Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
            End Using
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        Finally
            OnHeartbeat("Process Complete")
            SetObjectEnableDisable_ThreadSafe(btnStart, True)
            SetObjectEnableDisable_ThreadSafe(btnStop, False)
        End Try
    End Function
#End Region

#Region "CNC EOD"
    Private Async Function ViewDataCNCEODAsync() As Task
        Try
            Dim startDate As Date = GetDateTimePickerValue_ThreadSafe(dtpckrStartDate)
            Dim endDate As Date = GetDateTimePickerValue_ThreadSafe(dtpckrEndDate)
            Dim sourceData As Strategy.SourceOfData = Strategy.SourceOfData.None
            If GetRadioButtonChecked_ThreadSafe(rdbLive) Then
                sourceData = Strategy.SourceOfData.Live
            Else
                sourceData = Strategy.SourceOfData.Database
            End If
            Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
            Dim database As Common.DataBaseTable = Common.DataBaseTable.None
            Dim margin As Decimal = 0
            Dim tick As Decimal = 0
            Select Case stockType
                Case Trade.TypeOfStock.Cash
                    database = Common.DataBaseTable.EOD_Cash
                    margin = 1
                    tick = 0.05
                Case Trade.TypeOfStock.Commodity
                    database = Common.DataBaseTable.EOD_Commodity
                    margin = 1
                    tick = 1
                Case Trade.TypeOfStock.Currency
                    database = Common.DataBaseTable.EOD_Currency
                    margin = 1
                    tick = 0.0025
                Case Trade.TypeOfStock.Futures
                    database = Common.DataBaseTable.EOD_Futures
                    margin = 1
                    tick = 0.05
            End Select

#Region "HK Positional Strategy Rule"
            'For qntyTyp As Integer = 1 To 2
            '    For tgtMul As Decimal = 0.5 To 0.5
            '        Dim filename As String = Nothing
            '        Using backtestStrategy As New CNCEODGenericStrategy(canceller:=_canceller,
            '                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
            '                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
            '                                                    tradeStartTime:=TimeSpan.Parse("09:15:00"),
            '                                                    lastTradeEntryTime:=TimeSpan.Parse("15:29:59"),
            '                                                    eodExitTime:=TimeSpan.Parse("15:29:59"),
            '                                                    tickSize:=tick,
            '                                                    marginMultiplier:=margin,
            '                                                    timeframe:=1,
            '                                                    heikenAshiCandle:=False,
            '                                                    stockType:=stockType,
            '                                                    databaseTable:=database,
            '                                                    dataSource:=sourceData,
            '                                                    initialCapital:=Decimal.MaxValue / 2,
            '                                                    usableCapital:=Decimal.MaxValue / 2,
            '                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue / 2,
            '                                                    amountToBeWithdrawn:=50000)
            '            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

            '            With backtestStrategy
            '                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Vijay CNC Instrument Details.csv")
            '                .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Investment Stock List.csv")

            '                .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)

            '                Select Case .RuleNumber
            '                    Case 10
            '                        .RuleEntityData = New VijayCNCStrategyRule.StrategyRuleEntities With {.RefreshQuantityAtDayStart = False}
            '                    Case 18
            '                        .RuleEntityData = New InvestmentCNCStrategyRule.StrategyRuleEntities With
            '                            {.QuantityType = InvestmentCNCStrategyRule.TypeOfQuantity.AP}
            '                    Case 23
            '                        .RuleEntityData = New HKPositionalStrategyRule.StrategyRuleEntities With
            '                            {.QuantityType = qntyTyp,
            '                             .QuntityForLinear = 2,
            '                             .TargetMultiplier = tgtMul,
            '                             .TypeOfExit = HKPositionalStrategyRule.ExitType.CompoundingToMonthlyATR}

            '                        filename = String.Format("CNC Candle Capital {0},QuantityType {1},TargetMul {2}",
            '                                               If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital),
            '                                               CType(.RuleEntityData, HKPositionalStrategyRule.StrategyRuleEntities).QuantityType,
            '                                               CType(.RuleEntityData, HKPositionalStrategyRule.StrategyRuleEntities).TargetMultiplier)
            '                    Case 24
            '                        .RuleEntityData = New HKPositionalStrategyRule1.StrategyRuleEntities With
            '                            {.QuantityType = qntyTyp,
            '                             .QuntityForLinear = 2,
            '                             .TargetMultiplier = tgtMul,
            '                             .TypeOfExit = HKPositionalStrategyRule1.ExitType.CompoundingToMonthlyATR}

            '                        filename = String.Format("CNC Wick Capital {0},QuantityType {1},TargetMul {2}",
            '                                               If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital),
            '                                               CType(.RuleEntityData, HKPositionalStrategyRule1.StrategyRuleEntities).QuantityType,
            '                                               CType(.RuleEntityData, HKPositionalStrategyRule1.StrategyRuleEntities).TargetMultiplier)
            '                End Select

            '                .NumberOfTradeableStockPerDay = 10

            '                .NumberOfTradesPerDay = Integer.MaxValue
            '                .NumberOfTradesPerStockPerDay = Integer.MaxValue

            '                .TickBasedStrategy = True
            '            End With
            '            'Dim filename As String = String.Format("CNC Output Capital {3} {0}_{1}_{2}", Now.Hour, Now.Minute, Now.Second,
            '            '                                   If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital))
            '            Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
            '        End Using
            '    Next
            'Next
#End Region

#Region "SMI HK Positional Strategy Rule"
            'Using backtestStrategy As New CNCEODGenericStrategy(canceller:=_canceller,
            '                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
            '                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
            '                                                    tradeStartTime:=TimeSpan.Parse("09:15:00"),
            '                                                    lastTradeEntryTime:=TimeSpan.Parse("15:29:59"),
            '                                                    eodExitTime:=TimeSpan.Parse("15:29:59"),
            '                                                    tickSize:=tick,
            '                                                    marginMultiplier:=margin,
            '                                                    timeframe:=1,
            '                                                    heikenAshiCandle:=False,
            '                                                    stockType:=stockType,
            '                                                    databaseTable:=database,
            '                                                    dataSource:=sourceData,
            '                                                    initialCapital:=Decimal.MaxValue / 2,
            '                                                    usableCapital:=Decimal.MaxValue / 2,
            '                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue / 2,
            '                                                    amountToBeWithdrawn:=50000)
            '    AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

            '    With backtestStrategy
            '        '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Vijay CNC Instrument Details.csv")
            '        .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Investment Stock List.csv")

            '        .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)

            '        .RuleEntityData = New SMIHKPositionalStrategyRule.StrategyRuleEntities With
            '            {.QuantityType = SMIHKPositionalStrategyRule.TypeOfQuantity.Linear,
            '             .QuntityForLinear = 2}

            '        .NumberOfTradeableStockPerDay = 10

            '        .NumberOfTradesPerDay = Integer.MaxValue
            '        .NumberOfTradesPerStockPerDay = Integer.MaxValue

            '        .TickBasedStrategy = True
            '    End With
            '    Dim filename As String = String.Format("CNC SMI Capital {0},QuantityType {1}",
            '                                            If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital),
            '                                            CType(backtestStrategy.RuleEntityData, SMIHKPositionalStrategyRule.StrategyRuleEntities).QuantityType)
            '    Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
            'End Using
#End Region

#Region "ATR Positional Strategy Rule"
            ''Dim tgtMulList As List(Of Decimal) = New List(Of Decimal) From {0.5, 1, 2, 3}
            'Dim tgtMulList As List(Of Decimal) = New List(Of Decimal) From {2}
            ''Dim atrMulList As List(Of Decimal) = New List(Of Decimal) From {0.3, 0.5, 0.7, 0.9}
            'Dim atrMulList As List(Of Decimal) = New List(Of Decimal) From {0.3}
            'For Each atrMul In atrMulList
            '    'For qntyTyp As Integer = 1 To 2
            '    For qntyTyp As Integer = 2 To 2
            '        For Each tgtMul In tgtMulList
            '            If tgtMul < atrMul Then Continue For
            '            For prtlExt As Integer = 0 To 1
            '                Dim filename As String = Nothing
            '                Using backtestStrategy As New CNCEODGenericStrategy(canceller:=_canceller,
            '                                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
            '                                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
            '                                                                    tradeStartTime:=TimeSpan.Parse("09:15:00"),
            '                                                                    lastTradeEntryTime:=TimeSpan.Parse("15:29:59"),
            '                                                                    eodExitTime:=TimeSpan.Parse("15:29:59"),
            '                                                                    tickSize:=tick,
            '                                                                    marginMultiplier:=margin,
            '                                                                    timeframe:=1,
            '                                                                    heikenAshiCandle:=False,
            '                                                                    stockType:=stockType,
            '                                                                    databaseTable:=database,
            '                                                                    dataSource:=sourceData,
            '                                                                    initialCapital:=Decimal.MaxValue / 2,
            '                                                                    usableCapital:=Decimal.MaxValue / 2,
            '                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue / 2,
            '                                                                    amountToBeWithdrawn:=50000)
            '                    AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

            '                    With backtestStrategy
            '                        '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Vijay CNC Instrument Details.csv")
            '                        .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Investment Stock List.csv")

            '                        .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)

            '                        .RuleEntityData = New ATRPositionalStrategyRule.StrategyRuleEntities With
            '                            {.QuantityType = qntyTyp,
            '                             .TargetMultiplier = tgtMul,
            '                             .EntryATRMultiplier = atrMul,
            '                             .PartialExit = prtlExt}

            '                        filename = String.Format("CNC ATR Capital {0},QuantityType {1},TargetMul {2},Entry ATR Mul {3},Partial Exit {4}",
            '                                                If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital),
            '                                                CType(.RuleEntityData, ATRPositionalStrategyRule.StrategyRuleEntities).QuantityType,
            '                                                CType(.RuleEntityData, ATRPositionalStrategyRule.StrategyRuleEntities).TargetMultiplier,
            '                                                CType(.RuleEntityData, ATRPositionalStrategyRule.StrategyRuleEntities).EntryATRMultiplier,
            '                                                CType(.RuleEntityData, ATRPositionalStrategyRule.StrategyRuleEntities).PartialExit)

            '                        .NumberOfTradeableStockPerDay = 1

            '                        .NumberOfTradesPerDay = Integer.MaxValue
            '                        .NumberOfTradesPerStockPerDay = Integer.MaxValue

            '                        .TickBasedStrategy = True
            '                    End With
            '                    'Dim filename As String = String.Format("CNC Output Capital {3} {0}_{1}_{2}", Now.Hour, Now.Minute, Now.Second,
            '                    '                                   If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital))
            '                    Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
            '                End Using
            '            Next
            '        Next
            '    Next
            'Next
#End Region

#Region "Price Drop Positional Strategy Rule"
            'Dim tgtMulList As List(Of Decimal) = New List(Of Decimal) From {0.5, 1, 2, 3}
            Dim tgtMulList As List(Of Decimal) = New List(Of Decimal) From {1}
            'Dim atrMulList As List(Of Decimal) = New List(Of Decimal) From {0.3, 0.5, 0.7, 0.9}
            Dim prcDrpList As List(Of Decimal) = New List(Of Decimal) From {5}
            For Each prcDrp In prcDrpList
                'For qntyTyp As Integer = 1 To 2
                For qntyTyp As Integer = 2 To 2
                    For Each tgtMul In tgtMulList
                        'If tgtMul < atrMul Then Continue For
                        For prtlExt As Integer = 0 To 1
                            Dim filename As String = Nothing
                            Using backtestStrategy As New CNCEODGenericStrategy(canceller:=_canceller,
                                                                                exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                                exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                                tradeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                                lastTradeEntryTime:=TimeSpan.Parse("15:29:59"),
                                                                                eodExitTime:=TimeSpan.Parse("15:29:59"),
                                                                                tickSize:=tick,
                                                                                marginMultiplier:=margin,
                                                                                timeframe:=1,
                                                                                heikenAshiCandle:=False,
                                                                                stockType:=stockType,
                                                                                databaseTable:=database,
                                                                                dataSource:=sourceData,
                                                                                initialCapital:=Decimal.MaxValue / 2,
                                                                                usableCapital:=Decimal.MaxValue / 2,
                                                                                minimumEarnedCapitalToWithdraw:=Decimal.MaxValue / 2,
                                                                                amountToBeWithdrawn:=50000)
                                AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                                With backtestStrategy
                                    '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Vijay CNC Instrument Details.csv")
                                    .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Investment Stock List.csv")

                                    .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)

                                    .RuleEntityData = New PriceDropPositionalStrategyRule.StrategyRuleEntities With
                                        {.QuantityType = qntyTyp,
                                         .TargetMultiplier = tgtMul,
                                         .PriceDropPercentage = prcDrp,
                                         .PartialExit = prtlExt}

                                    filename = String.Format("CNC Price Drop Capital {0},QuantityType {1},TargetMul {2},Price Drop Percentage {3},Partial Exit {4}",
                                                            If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital),
                                                            CType(.RuleEntityData, PriceDropPositionalStrategyRule.StrategyRuleEntities).QuantityType,
                                                            CType(.RuleEntityData, PriceDropPositionalStrategyRule.StrategyRuleEntities).TargetMultiplier,
                                                            CType(.RuleEntityData, PriceDropPositionalStrategyRule.StrategyRuleEntities).PriceDropPercentage,
                                                            CType(.RuleEntityData, PriceDropPositionalStrategyRule.StrategyRuleEntities).PartialExit)

                                    .NumberOfTradeableStockPerDay = 1

                                    .NumberOfTradesPerDay = Integer.MaxValue
                                    .NumberOfTradesPerStockPerDay = Integer.MaxValue

                                    .TickBasedStrategy = True
                                End With
                                'Dim filename As String = String.Format("CNC Output Capital {3} {0}_{1}_{2}", Now.Hour, Now.Minute, Now.Second,
                                '                                   If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital))
                                Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                            End Using
                        Next
                    Next
                Next
            Next
#End Region

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        Finally
            OnHeartbeat("Process Complete")
            SetObjectEnableDisable_ThreadSafe(btnStart, True)
            SetObjectEnableDisable_ThreadSafe(btnStop, False)
        End Try
    End Function
#End Region

#Region "CNC Candle"
    Private Async Function ViewDataCNCCandleAsync() As Task
        Try
            Dim startDate As Date = GetDateTimePickerValue_ThreadSafe(dtpckrStartDate)
            Dim endDate As Date = GetDateTimePickerValue_ThreadSafe(dtpckrEndDate)
            Dim sourceData As Strategy.SourceOfData = Strategy.SourceOfData.None
            If GetRadioButtonChecked_ThreadSafe(rdbLive) Then
                sourceData = Strategy.SourceOfData.Live
            Else
                sourceData = Strategy.SourceOfData.Database
            End If
            Dim stockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
            Dim database As Common.DataBaseTable = Common.DataBaseTable.None
            Dim margin As Decimal = 0
            Dim tick As Decimal = 0
            Select Case stockType
                Case Trade.TypeOfStock.Cash
                    database = Common.DataBaseTable.Intraday_Cash
                    margin = 1
                    tick = 0.05
                Case Trade.TypeOfStock.Commodity
                    database = Common.DataBaseTable.Intraday_Commodity
                    margin = 1
                    tick = 1
                Case Trade.TypeOfStock.Currency
                    database = Common.DataBaseTable.Intraday_Currency
                    margin = 1
                    tick = 0.0025
                Case Trade.TypeOfStock.Futures
                    database = Common.DataBaseTable.Intraday_Futures
                    margin = 1
                    tick = 0.05
            End Select

            Dim tgtMulList As List(Of Decimal) = New List(Of Decimal) From {0.5, 1, 2, 3}
            For qntyTyp As Integer = 1 To 2
                For Each tgtMul In tgtMulList
                    Using backtestStrategy As New CNCCandleGenericStrategy(canceller:=_canceller,
                                                                    exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                    exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                    tradeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                    lastTradeEntryTime:=TimeSpan.Parse("15:29:59"),
                                                                    eodExitTime:=TimeSpan.Parse("15:29:59"),
                                                                    tickSize:=tick,
                                                                    marginMultiplier:=margin,
                                                                    timeframe:=180,
                                                                    heikenAshiCandle:=False,
                                                                    stockType:=stockType,
                                                                    databaseTable:=database,
                                                                    dataSource:=sourceData,
                                                                    initialCapital:=Decimal.MaxValue / 2,
                                                                    usableCapital:=Decimal.MaxValue / 2,
                                                                    minimumEarnedCapitalToWithdraw:=Decimal.MaxValue / 2,
                                                                    amountToBeWithdrawn:=5000)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Vijay CNC Instrument Details.csv")
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Investment Stock List.csv")

                            .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)

                            Select Case .RuleNumber
                                Case 26
                                    .RuleEntityData = New HKPositionalHourlyStrategyRule1.StrategyRuleEntities With
                                                {.QuantityType = qntyTyp,
                                                 .QuntityForLinear = 2,
                                                 .TargetMultiplier = tgtMul,
                                                 .TypeOfExit = HKPositionalHourlyStrategyRule1.ExitType.CompoundingToMonthlyATR}
                            End Select

                            .NumberOfTradeableStockPerDay = 10

                            .NumberOfTradesPerDay = Integer.MaxValue
                            .NumberOfTradesPerStockPerDay = Integer.MaxValue

                            .TickBasedStrategy = True
                        End With
                        Dim filename As String = String.Format("CNC Hourly Capital {0},QuantityType {1},TargetMul {2}",
                                                                If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital),
                                                                CType(backtestStrategy.RuleEntityData, HKPositionalHourlyStrategyRule1.StrategyRuleEntities).QuantityType,
                                                                CType(backtestStrategy.RuleEntityData, HKPositionalHourlyStrategyRule1.StrategyRuleEntities).TargetMultiplier)
                        Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                    End Using
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        Finally
            OnHeartbeat("Process Complete")
            SetObjectEnableDisable_ThreadSafe(btnStart, True)
            SetObjectEnableDisable_ThreadSafe(btnStop, False)
        End Try
    End Function
#End Region

    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        _canceller.Cancel()
    End Sub
End Class

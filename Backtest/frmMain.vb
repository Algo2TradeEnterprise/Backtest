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
        rdbCNC.Checked = My.Settings.CNC
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
        My.Settings.CNC = rdbCNC.Checked
        My.Settings.StartDate = dtpckrStartDate.Value
        My.Settings.EndDate = dtpckrEndDate.Value
        My.Settings.Save()

        If rdbMIS.Checked Then
            Await Task.Run(AddressOf ViewDataMISAsync).ConfigureAwait(False)
        ElseIf rdbCNC.Checked Then
            Await Task.Run(AddressOf ViewDataCNCAsync).ConfigureAwait(False)
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
                    margin = 70
                    tick = 0.0025
                Case Trade.TypeOfStock.Futures
                    database = Common.DataBaseTable.Intraday_Futures
                    margin = 30
                    tick = 0.05
            End Select

            For timeFrame As Integer = 1 To 1 Step 1
                For nmbrOfStock As Integer = 5 To 5 Step 1
                    For nmbrOfTradePerStock As Integer = 4 To 4 Step -1
                        For maxLossPercentagePerStock As Decimal = Decimal.MinValue To Decimal.MinValue Step 1
                            For tgtMul As Integer = 4 To 4 Step 1
                                For slMul As Integer = 1 To 1 Step 1
                                    For brkevenMvmnt As Integer = 0 To 0 Step 1
                                        For brkevenMul As Decimal = 2 / 3 To 2 / 3 Step 1
                                            For slMakeupTrade As Integer = Integer.MaxValue To Integer.MaxValue Step -1
                                                Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                                                  exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                                                  exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                                                  tradeStartTime:=TimeSpan.Parse("9:17:00"),
                                                                                                  lastTradeEntryTime:=TimeSpan.Parse("14:45:59"),
                                                                                                  eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                                                  tickSize:=tick,
                                                                                                  marginMultiplier:=margin,
                                                                                                  timeframe:=timeFrame,
                                                                                                  heikenAshiCandle:=False,
                                                                                                  stockType:=stockType,
                                                                                                  databaseTable:=database,
                                                                                                  dataSource:=sourceData,
                                                                                                  initialCapital:=Decimal.MaxValue / 2,
                                                                                                  usableCapital:=Decimal.MaxValue / 2,
                                                                                                  minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                                                                  amountToBeWithdrawn:=Decimal.MaxValue / 2)
                                                    AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                                                    With backtestStrategy
                                                        '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Future Stock List ATR Based.csv")
                                                        '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based Stocks.csv")
                                                        '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Pre Market Data.csv")
                                                        '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "BANKNIFTY.csv")
                                                        '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Vijay CNC Instrument Details.csv")
                                                        .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Volume spike Stock List.csv")

                                                        .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)
                                                        Select Case .RuleNumber
                                                            Case 1
                                                                .RuleEntityData = New HighVolumePinBarStrategyRule.StrategyRuleEntities With {.TargetMultiplier = tgtMul}
                                                            Case 2
                                                                .RuleEntityData = New MomentumReversalv2StrategyRule.StrategyRuleEntities With {.TargetMultiplier = tgtMul, .StoplossMultiplier = slMul, .BreakevenMovement = True, .ReEntryAtPreviousSignal = True}
                                                            Case 3
                                                                .RuleEntityData = New HighVolumePinBarv2StrategyRule.StrategyRuleEntities With {.TargetMultiplier = tgtMul, .StoplossMultiplier = slMul, .ModifyStoploss = True, .ReEntryAtPreviousSignal = True}
                                                            Case 4
                                                                .RuleEntityData = New DonchianFractalStrategyRule.StrategyRuleEntities With {.StoplossPercentage = slMul, .ModifyStoploss = True}
                                                            Case 5
                                                                .RuleEntityData = New SMIFractalStrategyRule.StrategyRuleEntities With {.ModifyStoploss = True}
                                                            Case 7
                                                                .RuleEntityData = New DayStartSMIStrategyRule.StrategyRuleEntities With {.TargetPercentage = tgtMul, .StoplossPercentage = slMul}
                                                            Case 10
                                                                .RuleEntityData = New VijayCNCStrategyRule.StrategyRuleEntities With {.RefreshQuantityAtDayStart = True}
                                                            Case 11
                                                                .RuleEntityData = New TIIOppositeBreakoutStrategyRule.StrategyRuleEntities With {.TargetMultiplier = tgtMul, .ModifyStoploss = True}
                                                            Case 12
                                                                .RuleEntityData = New FixedLevelBasedStrategyRule.StrategyRuleEntities With
                                                                {.TargetMultiplier = tgtMul,
                                                                .StoplossMultiplier = slMul,
                                                                .BreakevenMovement = brkevenMvmnt,
                                                                .BreakevenMultiplier = brkevenMul,
                                                                .LevelType = FixedLevelBasedStrategyRule.StrategyRuleEntities.TypeOfLevel.None,
                                                                .StoplossMakeupTrade = slMakeupTrade}
                                                        End Select


                                                        .NumberOfTradeableStockPerDay = nmbrOfStock

                                                        .NumberOfTradesPerDay = Integer.MaxValue
                                                        .NumberOfTradesPerStockPerDay = nmbrOfTradePerStock

                                                        .TrailingStoploss = False

                                                        .TickBasedStrategy = True

                                                        .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                                                        .StockMaxLossPercentagePerDay = maxLossPercentagePerStock

                                                        .ExitOnStockFixedTargetStoploss = False
                                                        .StockMaxProfitPerDay = Decimal.MaxValue
                                                        .StockMaxLossPerDay = Decimal.MinValue

                                                        .ExitOnOverAllFixedTargetStoploss = False
                                                        .OverAllProfitPerDay = Decimal.MaxValue
                                                        .OverAllLossPerDay = Decimal.MinValue

                                                        .TrailingMTM = False
                                                        .MTMSlab = 10000
                                                    End With
                                                    Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
                                                End Using
                                            Next
                                        Next
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.StackTrace, MsgBoxStyle.Critical)
        Finally
            OnHeartbeat("Process Complete")
            SetObjectEnableDisable_ThreadSafe(btnStart, True)
            SetObjectEnableDisable_ThreadSafe(btnStop, False)
        End Try
    End Function

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
                                                            timeframe:=1,
                                                            heikenAshiCandle:=False,
                                                            stockType:=stockType,
                                                            databaseTable:=database,
                                                            dataSource:=sourceData,
                                                            initialCapital:=Decimal.MaxValue / 2,
                                                            usableCapital:=Decimal.MaxValue / 2,
                                                            minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
                                                            amountToBeWithdrawn:=Decimal.MaxValue / 2)
                AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                With backtestStrategy
                    .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Vijay CNC Instrument Details.csv")

                    .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)

                    Select Case .RuleNumber
                        Case 10
                            .RuleEntityData = New VijayCNCStrategyRule.StrategyRuleEntities With {.RefreshQuantityAtDayStart = False}
                    End Select

                    .NumberOfTradeableStockPerDay = 1

                    .NumberOfTradesPerDay = Integer.MaxValue
                    .NumberOfTradesPerStockPerDay = Integer.MaxValue

                    .TickBasedStrategy = True
                End With
                Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
            End Using
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
End Class

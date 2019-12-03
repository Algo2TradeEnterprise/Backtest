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
            Await Task.Run(AddressOf ViewDataCNCEODAsync).ConfigureAwait(False)
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
                    margin = 98
                    tick = 0.0025
                Case Trade.TypeOfStock.Futures
                    database = Common.DataBaseTable.Intraday_Futures
                    margin = 30
                    tick = 0.05
            End Select

#Region "Dummy"
            'Dim tgtList As List(Of Decimal) = New List(Of Decimal) From {15}
            'For Each tgtMul As Decimal In tgtList
            '    For brkevn As Integer = 1 To 1
            '        For INRBsd As Integer = 1 To 1
            '            Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
            '                                                              exchangeStartTime:=TimeSpan.Parse("09:15:00"),
            '                                                              exchangeEndTime:=TimeSpan.Parse("15:29:59"),
            '                                                              tradeStartTime:=TimeSpan.Parse("9:16:00"),
            '                                                              lastTradeEntryTime:=TimeSpan.Parse("14:30:00"),
            '                                                              eodExitTime:=TimeSpan.Parse("15:15:00"),
            '                                                              tickSize:=tick,
            '                                                              marginMultiplier:=margin,
            '                                                              timeframe:=1,
            '                                                              heikenAshiCandle:=False,
            '                                                              stockType:=stockType,
            '                                                              databaseTable:=database,
            '                                                              dataSource:=sourceData,
            '                                                              initialCapital:=300000,
            '                                                              usableCapital:=300000,
            '                                                              minimumEarnedCapitalToWithdraw:=400000,
            '                                                              amountToBeWithdrawn:=100000)
            '                AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

            '                With backtestStrategy
            '                    '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Intraday Volume Spike for first 2 minute.csv")
            '                    .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "New ATR Based Stocks.csv")
            '                    '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "USDINR.csv")

            '                    .AllowBothDirectionEntryAtSameTime = False
            '                    .TrailingStoploss = False
            '                    .TickBasedStrategy = True
            '                    .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)
            '                    Select Case .RuleNumber
            '                        Case 1
            '                            .RuleEntityData = New HighVolumePinBarStrategyRule.StrategyRuleEntities With {.TargetMultiplier = 4}
            '                        Case 2
            '                            .RuleEntityData = New MomentumReversalv2StrategyRule.StrategyRuleEntities With {.TargetMultiplier = 4, .StoplossMultiplier = 1, .BreakevenMovement = True, .ReEntryAtPreviousSignal = True}
            '                        Case 3
            '                            .RuleEntityData = New HighVolumePinBarv2StrategyRule.StrategyRuleEntities With {.TargetMultiplier = 4, .StoplossMultiplier = 1, .ModifyStoploss = True, .ReEntryAtPreviousSignal = True}
            '                        Case 4
            '                            .RuleEntityData = New DonchianFractalStrategyRule.StrategyRuleEntities With {.StoplossPercentage = 1, .ModifyStoploss = True}
            '                        Case 5
            '                            .RuleEntityData = New SMIFractalStrategyRule.StrategyRuleEntities With {.ModifyStoploss = True}
            '                        Case 7
            '                            .RuleEntityData = New DayStartSMIStrategyRule.StrategyRuleEntities With {.TargetPercentage = 4, .StoplossPercentage = 1}
            '                        Case 10
            '                            .RuleEntityData = New VijayCNCStrategyRule.StrategyRuleEntities With {.RefreshQuantityAtDayStart = True}
            '                        Case 11
            '                            .RuleEntityData = New TIIOppositeBreakoutStrategyRule.StrategyRuleEntities With {.TargetMultiplier = 4, .ModifyStoploss = True}
            '                        Case 12
            '                            .RuleEntityData = New FixedLevelBasedStrategyRule.StrategyRuleEntities With
            '                            {.TargetMultiplier = 4,
            '                            .StoplossMultiplier = 1,
            '                            .BreakevenMovement = False,
            '                            .BreakevenMultiplier = 4,
            '                            .LevelType = FixedLevelBasedStrategyRule.StrategyRuleEntities.TypeOfLevel.None,
            '                            .StoplossMakeupTrade = False,
            '                            .MaxLossPercentageOfCapital = Decimal.MinValue,
            '                            .ModifyCandleTarget = True,
            '                            .ModifyNumberOfTrade = False}
            '                        Case 13
            '                            .RuleEntityData = New LowStoplossStrategyRule.StrategyRuleEntities With
            '                            {.StartingLevelMultiplier = 1,
            '                            .ChangeLevelAfterStoploss = False,
            '                            .AfterStoplossLevelMultiplier = 2,
            '                            .MaxStoploss = 1000,
            '                            .TargetMultiplier = 4,
            '                            .BreakevenMovement = True,
            '                            .ModifyNumberOfTrade = False,
            '                            .MaxPLToModifyNumberOfTrade = 0,
            '                            .MinimumCapital = 10000,
            '                            .MaxTargetPerTrade = Decimal.MaxValue,
            '                            .TypeOfSignal = LowStoplossStrategyRule.SignalType.DipInATR}
            '                        Case 14
            '                            .RuleEntityData = Nothing
            '                        Case 15
            '                            .RuleEntityData = New ReversalStrategyRule.StrategyRuleEntities With
            '                             {.TargetMultiplier = 3,
            '                              .NumberOfTradeOnNewSignal = 2,
            '                              .BreakevenMovement = True}
            '                        Case 16
            '                            .RuleEntityData = New PinbarBreakoutStrategyRule.StrategyRuleEntities With
            '                             {.MinimumInvestmentPerStock = 15000,
            '                              .MaxLossPerTradeMultiplier = 0.5,
            '                              .MinLossPercentagePerTrade = 0.1,
            '                              .PinbarTailPercentage = 50,
            '                              .TargetMultiplier = 1,
            '                              .BreakevenMovement = True,
            '                              .SignalAtDayHighLow = False,
            '                              .StopAtFirstTarget = False}
            '                        Case 17
            '                            .RuleEntityData = New LowSLPinbarStrategyRule.StrategyRuleEntities With
            '                             {.MinimumInvestmentPerStock = 15000,
            '                              .MaxLossPerTrade = -1000,
            '                              .PinbarTailPercentage = 65,
            '                              .TargetMultiplier = 2,
            '                              .BreakevenMovement = True,
            '                              .StopAtFirstTarget = False,
            '                              .AllowMomentumReversal = False}
            '                        Case 19
            '                            .RuleEntityData = New LowStoplossWickStrategyRule.StrategyRuleEntities With
            '                                {.MinimumInvestmentPerStock = 15000,
            '                                 .MinStoploss = 700,
            '                                 .MaxStoploss = 1500,
            '                                 .TargetMultiplier = 2,
            '                                 .MinimumStockMaxExitPerTrade = True,
            '                                 .TypeOfSLMakeup = LowStoplossWickStrategyRule.StoplossMakeupType.SingleLossMakeup
            '                                }
            '                        Case 20
            '                            .RuleEntityData = New LowStoplossCandleStrategyRule.StrategyRuleEntities With
            '                                {.MinimumInvestmentPerStock = 15000,
            '                                 .MinStoploss = 700,
            '                                 .MaxStoploss = 1500,
            '                                 .TargetMultiplier = 2,
            '                                 .MinimumStockMaxExitPerTrade = True,
            '                                 .TypeOfSLMakeup = LowStoplossCandleStrategyRule.StoplossMakeupType.SingleLossMakeup
            '                                }
            '                        Case 21
            '                            .AllowBothDirectionEntryAtSameTime = True
            '                            .RuleEntityData = New PairTradingStrategyRule.StrategyRuleEntities With
            '                                {.TargetMultiplier = tgtMul,
            '                                 .BreakevenMovement = brkevn,
            '                                 .INRBasedTarget = INRBsd,
            '                                 .MinumumInvesment = 15000
            '                                }
            '                        Case 22
            '                            .RuleEntityData = New CoinFlipAtResistanceStrategyRule.StrategyRuleEntities With
            '                                {.TargetMultiplier = tgtMul,
            '                                 .BreakevenMovement = brkevn,
            '                                 .INRBasedTarget = INRBsd,
            '                                 .MinumumInvesment = 15000
            '                                }
            '                    End Select


            '                    .NumberOfTradeableStockPerDay = 10

            '                    .NumberOfTradesPerStockPerDay = Integer.MaxValue

            '                    .StockMaxProfitPercentagePerDay = Decimal.MaxValue
            '                    .StockMaxLossPercentagePerDay = Decimal.MinValue

            '                    .ExitOnStockFixedTargetStoploss = False
            '                    .StockMaxProfitPerDay = Decimal.MaxValue
            '                    .StockMaxLossPerDay = Decimal.MinValue

            '                    .ExitOnOverAllFixedTargetStoploss = False
            '                    .OverAllProfitPerDay = Decimal.MaxValue
            '                    .OverAllLossPerDay = Decimal.MinValue

            '                    .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
            '                    .MTMSlab = Math.Abs(.OverAllLossPerDay)
            '                    .MovementSlab = .MTMSlab / 2
            '                    .RealtimeTrailingPercentage = 50
            '                End With

            '                'Dim ruleData As LowStoplossCandleStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
            '                'Dim filename As String = String.Format("TF {0},StkMxPft {1},StkMxLs {2},OvrAlPft {3},OvrAlLs {4},TrlMTMTyp {5},ExtPTrd {6},SLMkupTyp {7}",
            '                '                                       backtestStrategy.SignalTimeFrame,
            '                '                                       If(backtestStrategy.StockMaxProfitPerDay <> Decimal.MaxValue, backtestStrategy.StockMaxProfitPerDay, "∞"),
            '                '                                       If(backtestStrategy.StockMaxLossPerDay <> Decimal.MinValue, backtestStrategy.StockMaxLossPerDay, "∞"),
            '                '                                       If(backtestStrategy.OverAllProfitPerDay <> Decimal.MaxValue, backtestStrategy.OverAllProfitPerDay, "∞"),
            '                '                                       If(backtestStrategy.OverAllLossPerDay <> Decimal.MinValue, backtestStrategy.OverAllLossPerDay, "∞"),
            '                '                                       backtestStrategy.TypeOfMTMTrailing.ToString,
            '                '                                       ruleData.MinimumStockMaxExitPerTrade,
            '                '                                       ruleData.TypeOfSLMakeup.ToString)

            '                Dim ruleData As CoinFlipAtResistanceStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
            '                Dim filename As String = String.Format("TF {0},StkMxPft {1},StkMxLs {2},OvrAlPft {3},OvrAlLs {4},TgtMul {5},Brkevn {6},INRBsd {7}",
            '                                                       backtestStrategy.SignalTimeFrame,
            '                                                       If(backtestStrategy.StockMaxProfitPerDay <> Decimal.MaxValue, backtestStrategy.StockMaxProfitPerDay, "∞"),
            '                                                       If(backtestStrategy.StockMaxLossPerDay <> Decimal.MinValue, backtestStrategy.StockMaxLossPerDay, "∞"),
            '                                                       If(backtestStrategy.OverAllProfitPerDay <> Decimal.MaxValue, backtestStrategy.OverAllProfitPerDay, "∞"),
            '                                                       If(backtestStrategy.OverAllLossPerDay <> Decimal.MinValue, backtestStrategy.OverAllLossPerDay, "∞"),
            '                                                       ruleData.TargetMultiplier,
            '                                                       ruleData.BreakevenMovement,
            '                                                       ruleData.INRBasedTarget)

            '                Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
            '            End Using
            '        Next
            '    Next
            'Next
#End Region

#Region "Low SL"
            'For signalType As Integer = 17 To 17 Step 1
            '    For strtLvlMul As Integer = 1 To 1 Step 1
            '        For chngLvlAfterSl As Integer = 0 To 0 Step 1
            '            For aftrSLLvlMul As Integer = 2 To 2 Step 1
            '                For nmbrOfTradePerStock As Integer = 20 To 20 Step 1
            '                    For mdfyNmbrOfTrd As Integer = 0 To 0 Step 1
            '                        For overallMaxLoss As Decimal = Decimal.MinValue To Decimal.MinValue Step 1
            '                            For brkEvnMvmnt As Integer = 1 To 1 Step 1
            '                                For tradeMaxProfit As Decimal = Decimal.MaxValue To Decimal.MaxValue Step -1
            '                                    If brkEvnMvmnt = 1 AndAlso tradeMaxProfit <> Decimal.MaxValue Then Continue For
            '                                    For stockMaxProfit As Decimal = Decimal.MaxValue To Decimal.MaxValue Step -1
            '                                        If brkEvnMvmnt = 1 AndAlso stockMaxProfit <> Decimal.MaxValue Then Continue For
            '                                        If strtLvlMul = 2 AndAlso chngLvlAfterSl = 1 Then Continue For
            '                                        Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
            '                                                                                          exchangeStartTime:=TimeSpan.Parse("09:15:00"),
            '                                                                                          exchangeEndTime:=TimeSpan.Parse("15:29:59"),
            '                                                                                          tradeStartTime:=TimeSpan.Parse("9:16:00"),
            '                                                                                          lastTradeEntryTime:=TimeSpan.Parse("14:45:59"),
            '                                                                                          eodExitTime:=TimeSpan.Parse("15:15:00"),
            '                                                                                          tickSize:=tick,
            '                                                                                          marginMultiplier:=margin,
            '                                                                                          timeframe:=1,
            '                                                                                          heikenAshiCandle:=False,
            '                                                                                          stockType:=stockType,
            '                                                                                          databaseTable:=database,
            '                                                                                          dataSource:=sourceData,
            '                                                                                          initialCapital:=Decimal.MaxValue / 2,
            '                                                                                          usableCapital:=Decimal.MaxValue / 2,
            '                                                                                          minimumEarnedCapitalToWithdraw:=Decimal.MaxValue,
            '                                                                                          amountToBeWithdrawn:=Decimal.MaxValue / 2)
            '                                            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

            '                                            With backtestStrategy
            '                                                .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based Stocks.csv")
            '                                                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Pre Market Data.csv")
            '                                                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "BANKNIFTY.csv")
            '                                                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Vijay CNC Instrument Details.csv")
            '                                                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Volume spike Stock List with abs ATR.csv")

            '                                                .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)
            '                                                Select Case .RuleNumber
            '                                                    Case 1
            '                                                        .RuleEntityData = New HighVolumePinBarStrategyRule.StrategyRuleEntities With {.TargetMultiplier = 4}
            '                                                    Case 2
            '                                                        .RuleEntityData = New MomentumReversalv2StrategyRule.StrategyRuleEntities With {.TargetMultiplier = 4, .StoplossMultiplier = 1, .BreakevenMovement = True, .ReEntryAtPreviousSignal = True}
            '                                                    Case 3
            '                                                        .RuleEntityData = New HighVolumePinBarv2StrategyRule.StrategyRuleEntities With {.TargetMultiplier = 4, .StoplossMultiplier = 1, .ModifyStoploss = True, .ReEntryAtPreviousSignal = True}
            '                                                    Case 4
            '                                                        .RuleEntityData = New DonchianFractalStrategyRule.StrategyRuleEntities With {.StoplossPercentage = 1, .ModifyStoploss = True}
            '                                                    Case 5
            '                                                        .RuleEntityData = New SMIFractalStrategyRule.StrategyRuleEntities With {.ModifyStoploss = True}
            '                                                    Case 7
            '                                                        .RuleEntityData = New DayStartSMIStrategyRule.StrategyRuleEntities With {.TargetPercentage = 4, .StoplossPercentage = 1}
            '                                                    Case 10
            '                                                        .RuleEntityData = New VijayCNCStrategyRule.StrategyRuleEntities With {.RefreshQuantityAtDayStart = True}
            '                                                    Case 11
            '                                                        .RuleEntityData = New TIIOppositeBreakoutStrategyRule.StrategyRuleEntities With {.TargetMultiplier = 4, .ModifyStoploss = True}
            '                                                    Case 12
            '                                                        .RuleEntityData = New FixedLevelBasedStrategyRule.StrategyRuleEntities With
            '                                                        {.TargetMultiplier = 4,
            '                                                        .StoplossMultiplier = 1,
            '                                                        .BreakevenMovement = False,
            '                                                        .BreakevenMultiplier = 4,
            '                                                        .LevelType = FixedLevelBasedStrategyRule.StrategyRuleEntities.TypeOfLevel.None,
            '                                                        .StoplossMakeupTrade = False,
            '                                                        .MaxLossPercentageOfCapital = Decimal.MinValue,
            '                                                        .ModifyCandleTarget = True,
            '                                                        .ModifyNumberOfTrade = False}
            '                                                    Case 13
            '                                                        .RuleEntityData = New LowStoplossStrategyRule.StrategyRuleEntities With
            '                                                        {.StartingLevelMultiplier = strtLvlMul,
            '                                                        .ChangeLevelAfterStoploss = chngLvlAfterSl,
            '                                                        .AfterStoplossLevelMultiplier = aftrSLLvlMul,
            '                                                        .MaxStoploss = 1000,
            '                                                        .TargetMultiplier = 4,
            '                                                        .BreakevenMovement = brkEvnMvmnt,
            '                                                        .ModifyNumberOfTrade = mdfyNmbrOfTrd,
            '                                                        .MaxPLToModifyNumberOfTrade = 0,
            '                                                        .MinimumCapital = 10000,
            '                                                        .MaxTargetPerTrade = tradeMaxProfit,
            '                                                        .TypeOfSignal = signalType}
            '                                                End Select


            '                                                .NumberOfTradeableStockPerDay = 20

            '                                                .NumberOfTradesPerDay = Integer.MaxValue
            '                                                .NumberOfTradesPerStockPerDay = nmbrOfTradePerStock

            '                                                .TrailingStoploss = False

            '                                                .TickBasedStrategy = True

            '                                                .StockMaxProfitPercentagePerDay = Decimal.MaxValue
            '                                                .StockMaxLossPercentagePerDay = Decimal.MinValue

            '                                                .ExitOnStockFixedTargetStoploss = True
            '                                                .StockMaxProfitPerDay = stockMaxProfit
            '                                                .StockMaxLossPerDay = Decimal.MinValue

            '                                                .ExitOnOverAllFixedTargetStoploss = True
            '                                                .OverAllProfitPerDay = Decimal.MaxValue
            '                                                .OverAllLossPerDay = overallMaxLoss

            '                                                .TrailingMTM = False
            '                                                .MTMSlab = 10000
            '                                            End With
            '                                            Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
            '                                        End Using
            '                                    Next
            '                                Next
            '                            Next
            '                        Next
            '                    Next
            '                Next
            '            Next
            '        Next
            '    Next
            'Next
#End Region

#Region "Pinbar Breakout"
            ''Dim stkPrftList As List(Of Decimal) = New List(Of Decimal) From {1, Decimal.MaxValue}
            'Dim stkPrftList As List(Of Decimal) = New List(Of Decimal) From {Decimal.MaxValue}
            'For trdAtDayHL As Integer = 0 To 0 Step 1
            '    For trgtMul As Decimal = 1 To 1 Step 1
            '        For brkevnMvmnt As Integer = 1 To 1 Step 1
            '            For stopAtFirstTarget As Integer = 0 To 0 Step 1
            '                For stockMaxLossMultipler As Decimal = 2 To 2 Step 1
            '                    For Each stockMaxProfitMultipler As Decimal In stkPrftList
            '                        If stockMaxProfitMultipler <> Decimal.MaxValue AndAlso stopAtFirstTarget = 1 Then Continue For
            '                        Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
            '                                                                          exchangeStartTime:=TimeSpan.Parse("09:15:00"),
            '                                                                          exchangeEndTime:=TimeSpan.Parse("15:29:59"),
            '                                                                          tradeStartTime:=TimeSpan.Parse("9:20:00"),
            '                                                                          lastTradeEntryTime:=TimeSpan.Parse("14:45:59"),
            '                                                                          eodExitTime:=TimeSpan.Parse("15:15:00"),
            '                                                                          tickSize:=tick,
            '                                                                          marginMultiplier:=margin,
            '                                                                          timeframe:=5,
            '                                                                          heikenAshiCandle:=False,
            '                                                                          stockType:=stockType,
            '                                                                          databaseTable:=database,
            '                                                                          dataSource:=sourceData,
            '                                                                          initialCapital:=250000,
            '                                                                          usableCapital:=200000,
            '                                                                          minimumEarnedCapitalToWithdraw:=300000,
            '                                                                          amountToBeWithdrawn:=100000)
            '                            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

            '                            With backtestStrategy
            '                                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "ATR Based Stocks.csv")
            '                                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Pre Market Data.csv")
            '                                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "BANKNIFTY.csv")
            '                                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Vijay CNC Instrument Details.csv")
            '                                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Volume spike Stock List with abs ATR.csv")
            '                                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Nifty 50.csv")
            '                                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Multi Target ATR Based Stocks.csv")
            '                                '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Pinbar Stocklist.csv")
            '                                .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Low SL Pinbar Stocklist.csv")

            '                                .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)
            '                                Select Case .RuleNumber
            '                                    Case 1
            '                                        .RuleEntityData = New HighVolumePinBarStrategyRule.StrategyRuleEntities With {.TargetMultiplier = 4}
            '                                    Case 2
            '                                        .RuleEntityData = New MomentumReversalv2StrategyRule.StrategyRuleEntities With {.TargetMultiplier = 4, .StoplossMultiplier = 1, .BreakevenMovement = True, .ReEntryAtPreviousSignal = True}
            '                                    Case 3
            '                                        .RuleEntityData = New HighVolumePinBarv2StrategyRule.StrategyRuleEntities With {.TargetMultiplier = 4, .StoplossMultiplier = 1, .ModifyStoploss = True, .ReEntryAtPreviousSignal = True}
            '                                    Case 4
            '                                        .RuleEntityData = New DonchianFractalStrategyRule.StrategyRuleEntities With {.StoplossPercentage = 1, .ModifyStoploss = True}
            '                                    Case 5
            '                                        .RuleEntityData = New SMIFractalStrategyRule.StrategyRuleEntities With {.ModifyStoploss = True}
            '                                    Case 7
            '                                        .RuleEntityData = New DayStartSMIStrategyRule.StrategyRuleEntities With {.TargetPercentage = 4, .StoplossPercentage = 1}
            '                                    Case 10
            '                                        .RuleEntityData = New VijayCNCStrategyRule.StrategyRuleEntities With {.RefreshQuantityAtDayStart = True}
            '                                    Case 11
            '                                        .RuleEntityData = New TIIOppositeBreakoutStrategyRule.StrategyRuleEntities With {.TargetMultiplier = 4, .ModifyStoploss = True}
            '                                    Case 12
            '                                        .RuleEntityData = New FixedLevelBasedStrategyRule.StrategyRuleEntities With
            '                                        {.TargetMultiplier = 4,
            '                                        .StoplossMultiplier = 1,
            '                                        .BreakevenMovement = False,
            '                                        .BreakevenMultiplier = 4,
            '                                        .LevelType = FixedLevelBasedStrategyRule.StrategyRuleEntities.TypeOfLevel.None,
            '                                        .StoplossMakeupTrade = False,
            '                                        .MaxLossPercentageOfCapital = Decimal.MinValue,
            '                                        .ModifyCandleTarget = True,
            '                                        .ModifyNumberOfTrade = False}
            '                                    Case 13
            '                                        .RuleEntityData = New LowStoplossStrategyRule.StrategyRuleEntities With
            '                                        {.StartingLevelMultiplier = 1,
            '                                        .ChangeLevelAfterStoploss = False,
            '                                        .AfterStoplossLevelMultiplier = 2,
            '                                        .MaxStoploss = 1000,
            '                                        .TargetMultiplier = 4,
            '                                        .BreakevenMovement = True,
            '                                        .ModifyNumberOfTrade = False,
            '                                        .MaxPLToModifyNumberOfTrade = 0,
            '                                        .MinimumCapital = 10000,
            '                                        .MaxTargetPerTrade = Decimal.MaxValue,
            '                                        .TypeOfSignal = LowStoplossStrategyRule.SignalType.DipInATR}
            '                                    Case 14
            '                                        .RuleEntityData = Nothing
            '                                    Case 15
            '                                        .RuleEntityData = New ReversalStrategyRule.StrategyRuleEntities With
            '                                         {.TargetMultiplier = 3,
            '                                          .NumberOfTradeOnNewSignal = 2,
            '                                          .BreakevenMovement = True}
            '                                    Case 16
            '                                        .RuleEntityData = New PinbarBreakoutStrategyRule.StrategyRuleEntities With
            '                                         {.MinimumInvestmentPerStock = 15000,
            '                                          .MaxLossPerTradeMultiplier = 0.5,
            '                                          .MinLossPercentagePerTrade = 0.1,
            '                                          .PinbarTailPercentage = 50,
            '                                          .TargetMultiplier = trgtMul,
            '                                          .BreakevenMovement = brkevnMvmnt,
            '                                          .SignalAtDayHighLow = trdAtDayHL,
            '                                          .StopAtFirstTarget = stopAtFirstTarget}
            '                                    Case 17
            '                                        .RuleEntityData = New LowSLPinbarStrategyRule.StrategyRuleEntities With
            '                                         {.MinimumInvestmentPerStock = 15000,
            '                                          .MaxLossPerTrade = -1000,
            '                                          .PinbarTailPercentage = 70,
            '                                          .TargetMultiplier = 2,
            '                                          .BreakevenMovement = True,
            '                                          .StopAtFirstTarget = False}
            '                                End Select


            '                                .NumberOfTradeableStockPerDay = Integer.MaxValue

            '                                .NumberOfTradesPerDay = Integer.MaxValue
            '                                .NumberOfTradesPerStockPerDay = Integer.MaxValue

            '                                .TrailingStoploss = False

            '                                .TickBasedStrategy = False

            '                                .StockMaxProfitPercentagePerDay = stockMaxProfitMultipler
            '                                .StockMaxLossPercentagePerDay = stockMaxLossMultipler

            '                                .ExitOnStockFixedTargetStoploss = False
            '                                .StockMaxProfitPerDay = Decimal.MaxValue
            '                                .StockMaxLossPerDay = Decimal.MinValue

            '                                .ExitOnOverAllFixedTargetStoploss = False
            '                                .OverAllProfitPerDay = Decimal.MaxValue
            '                                .OverAllLossPerDay = Decimal.MinValue

            '                                .TrailingMTM = False
            '                                .MTMSlab = 20000
            '                            End With
            '                            Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
            '                        End Using
            '                    Next
            '                Next
            '            Next
            '        Next
            '    Next
            'Next
#End Region

#Region "Pair Trading & Coin Flip"
            'Dim tgtList As List(Of Decimal) = New List(Of Decimal) From {15}
            'For Each tgtMul As Decimal In tgtList
            '    For brkevn As Integer = 1 To 1
            '        For INRBsd As Integer = 1 To 1
            '            Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
            '                                                              exchangeStartTime:=TimeSpan.Parse("09:15:00"),
            '                                                              exchangeEndTime:=TimeSpan.Parse("15:29:59"),
            '                                                              tradeStartTime:=TimeSpan.Parse("9:16:00"),
            '                                                              lastTradeEntryTime:=TimeSpan.Parse("14:30:00"),
            '                                                              eodExitTime:=TimeSpan.Parse("15:15:00"),
            '                                                              tickSize:=tick,
            '                                                              marginMultiplier:=margin,
            '                                                              timeframe:=1,
            '                                                              heikenAshiCandle:=False,
            '                                                              stockType:=stockType,
            '                                                              databaseTable:=database,
            '                                                              dataSource:=sourceData,
            '                                                              initialCapital:=300000,
            '                                                              usableCapital:=300000,
            '                                                              minimumEarnedCapitalToWithdraw:=400000,
            '                                                              amountToBeWithdrawn:=100000)
            '                AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

            '                With backtestStrategy
            '                    '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Intraday Volume Spike for first 2 minute.csv")
            '                    .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "New ATR Based Stocks.csv")
            '                    '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "USDINR.csv")

            '                    .AllowBothDirectionEntryAtSameTime = False
            '                    .TrailingStoploss = False
            '                    .TickBasedStrategy = True
            '                    .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)
            '                    Select Case .RuleNumber
            '                        Case 21
            '                            .AllowBothDirectionEntryAtSameTime = True
            '                            .RuleEntityData = New PairTradingStrategyRule.StrategyRuleEntities With
            '                                {.TargetMultiplier = tgtMul,
            '                                 .BreakevenMovement = brkevn,
            '                                 .INRBasedTarget = INRBsd,
            '                                 .MinumumInvesment = 15000
            '                                }
            '                        Case 22
            '                            .RuleEntityData = New CoinFlipAtResistanceStrategyRule.StrategyRuleEntities With
            '                                {.TargetMultiplier = tgtMul,
            '                                 .BreakevenMovement = brkevn,
            '                                 .INRBasedTarget = INRBsd,
            '                                 .MinumumInvesment = 15000
            '                                }
            '                    End Select


            '                    .NumberOfTradeableStockPerDay = 10

            '                    .NumberOfTradesPerStockPerDay = Integer.MaxValue

            '                    .StockMaxProfitPercentagePerDay = Decimal.MaxValue
            '                    .StockMaxLossPercentagePerDay = Decimal.MinValue

            '                    .ExitOnStockFixedTargetStoploss = False
            '                    .StockMaxProfitPerDay = Decimal.MaxValue
            '                    .StockMaxLossPerDay = Decimal.MinValue

            '                    .ExitOnOverAllFixedTargetStoploss = False
            '                    .OverAllProfitPerDay = Decimal.MaxValue
            '                    .OverAllLossPerDay = Decimal.MinValue

            '                    .TypeOfMTMTrailing = Strategy.MTMTrailingType.None
            '                    .MTMSlab = Math.Abs(.OverAllLossPerDay)
            '                    .MovementSlab = .MTMSlab / 2
            '                    .RealtimeTrailingPercentage = 50
            '                End With

            '                Dim ruleData As CoinFlipAtResistanceStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
            '                Dim filename As String = String.Format("TF {0},StkMxPft {1},StkMxLs {2},OvrAlPft {3},OvrAlLs {4},TgtMul {5},Brkevn {6},INRBsd {7}",
            '                                                       backtestStrategy.SignalTimeFrame,
            '                                                       If(backtestStrategy.StockMaxProfitPerDay <> Decimal.MaxValue, backtestStrategy.StockMaxProfitPerDay, "∞"),
            '                                                       If(backtestStrategy.StockMaxLossPerDay <> Decimal.MinValue, backtestStrategy.StockMaxLossPerDay, "∞"),
            '                                                       If(backtestStrategy.OverAllProfitPerDay <> Decimal.MaxValue, backtestStrategy.OverAllProfitPerDay, "∞"),
            '                                                       If(backtestStrategy.OverAllLossPerDay <> Decimal.MinValue, backtestStrategy.OverAllLossPerDay, "∞"),
            '                                                       ruleData.TargetMultiplier,
            '                                                       ruleData.BreakevenMovement,
            '                                                       ruleData.INRBasedTarget)

            '                Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
            '            End Using
            '        Next
            '    Next
            'Next
#End Region

#Region "Low Stoploss Wick & Candle"
            For ovrAlLoss As Decimal = -10000 To -10000 Step 1000
                For stkMaxLoss As Decimal = -3000 To -3000 Step 1000
                    For trlngMTM As Integer = 2 To 4
                        For slMkupType As Integer = 1 To 2
                            Using backtestStrategy As New MISGenericStrategy(canceller:=_canceller,
                                                                              exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                              exchangeEndTime:=TimeSpan.Parse("15:29:59"),
                                                                              tradeStartTime:=TimeSpan.Parse("9:17:00"),
                                                                              lastTradeEntryTime:=TimeSpan.Parse("14:40:59"),
                                                                              eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                              tickSize:=tick,
                                                                              marginMultiplier:=margin,
                                                                              timeframe:=1,
                                                                              heikenAshiCandle:=False,
                                                                              stockType:=stockType,
                                                                              databaseTable:=database,
                                                                              dataSource:=sourceData,
                                                                              initialCapital:=300000,
                                                                              usableCapital:=300000,
                                                                              minimumEarnedCapitalToWithdraw:=400000,
                                                                              amountToBeWithdrawn:=100000)
                                AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                                With backtestStrategy
                                    .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Intraday Volume Spike for first 2 minute.csv")
                                    '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "New ATR Based Stocks.csv")

                                    .AllowBothDirectionEntryAtSameTime = False
                                    .TrailingStoploss = False
                                    .TickBasedStrategy = True
                                    .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)
                                    Select Case .RuleNumber
                                        Case 19
                                            .TickBasedStrategy = False
                                            .RuleEntityData = New LowStoplossWickStrategyRule.StrategyRuleEntities With
                                                {.MinimumInvestmentPerStock = 15000,
                                                 .MinStoploss = 700,
                                                 .MaxStoploss = 1500,
                                                 .TargetMultiplier = 2,
                                                 .MinimumStockMaxExitPerTrade = True,
                                                 .TypeOfSLMakeup = slMkupType
                                                }
                                        Case 20
                                            .RuleEntityData = New LowStoplossCandleStrategyRule.StrategyRuleEntities With
                                                {.MinimumInvestmentPerStock = 15000,
                                                 .MinStoploss = 700,
                                                 .MaxStoploss = 1500,
                                                 .TargetMultiplier = 2,
                                                 .MinimumStockMaxExitPerTrade = True,
                                                 .TypeOfSLMakeup = slMkupType
                                                }
                                    End Select

                                    .NumberOfTradeableStockPerDay = 10

                                    .NumberOfTradesPerStockPerDay = Integer.MaxValue

                                    .StockMaxProfitPercentagePerDay = Decimal.MaxValue
                                    .StockMaxLossPercentagePerDay = Decimal.MinValue

                                    .ExitOnStockFixedTargetStoploss = True
                                    .StockMaxProfitPerDay = 4000
                                    .StockMaxLossPerDay = stkMaxLoss

                                    .ExitOnOverAllFixedTargetStoploss = True
                                    .OverAllProfitPerDay = Decimal.MaxValue
                                    .OverAllLossPerDay = ovrAlLoss

                                    .TypeOfMTMTrailing = trlngMTM
                                    .MTMSlab = Math.Abs(.OverAllLossPerDay)
                                    .MovementSlab = .MTMSlab / 2
                                    .RealtimeTrailingPercentage = 50
                                End With

                                Dim ruleData As LowStoplossWickStrategyRule.StrategyRuleEntities = backtestStrategy.RuleEntityData
                                Dim filename As String = String.Format("TF {0},StkMxPft {1},StkMxLs {2},OvrAlPft {3},OvrAlLs {4},TrlMTMTyp {5},ExtPTrd {6},SLMkupTyp {7}",
                                                                       backtestStrategy.SignalTimeFrame,
                                                                       If(backtestStrategy.StockMaxProfitPerDay <> Decimal.MaxValue, backtestStrategy.StockMaxProfitPerDay, "∞"),
                                                                       If(backtestStrategy.StockMaxLossPerDay <> Decimal.MinValue, backtestStrategy.StockMaxLossPerDay, "∞"),
                                                                       If(backtestStrategy.OverAllProfitPerDay <> Decimal.MaxValue, backtestStrategy.OverAllProfitPerDay, "∞"),
                                                                       If(backtestStrategy.OverAllLossPerDay <> Decimal.MinValue, backtestStrategy.OverAllLossPerDay, "∞"),
                                                                       backtestStrategy.TypeOfMTMTrailing.ToString,
                                                                       ruleData.MinimumStockMaxExitPerTrade,
                                                                       ruleData.TypeOfSLMakeup.ToString)

                                Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
                            End Using
                        Next
                    Next
                Next
            Next
#End Region

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
                                                            initialCapital:=20000,
                                                            usableCapital:=15000,
                                                            minimumEarnedCapitalToWithdraw:=25000,
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
            For qntyTyp As Integer = 2 To 2
                For tgtMul As Decimal = 1 To 1
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
                                                                initialCapital:=475000,
                                                                usableCapital:=475000,
                                                                minimumEarnedCapitalToWithdraw:=Decimal.MaxValue / 2,
                                                                amountToBeWithdrawn:=50000)
                        AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                        With backtestStrategy
                            '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Vijay CNC Instrument Details.csv")
                            .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Investment Stock List.csv")

                            .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)

                            Select Case .RuleNumber
                                Case 10
                                    .RuleEntityData = New VijayCNCStrategyRule.StrategyRuleEntities With {.RefreshQuantityAtDayStart = False}
                                Case 18
                                    .RuleEntityData = New InvestmentCNCStrategyRule.StrategyRuleEntities With
                                        {.QuantityType = InvestmentCNCStrategyRule.TypeOfQuantity.AP}
                                Case 23
                                    .RuleEntityData = New HKPositionalStrategyRule.StrategyRuleEntities With
                                        {.QuantityType = qntyTyp,
                                         .QuntityForLinear = 2,
                                         .TargetMultiplier = tgtMul,
                                         .TypeOfExit = HKPositionalStrategyRule.ExitType.CompoundingToMonthlyATR}

                                    filename = String.Format("CNC Candle Capital {0},QuantityType {1},TargetMul {2}",
                                                           If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital),
                                                           CType(.RuleEntityData, HKPositionalStrategyRule.StrategyRuleEntities).QuantityType,
                                                           CType(.RuleEntityData, HKPositionalStrategyRule.StrategyRuleEntities).TargetMultiplier)
                                Case 24
                                    .RuleEntityData = New HKPositionalStrategyRule1.StrategyRuleEntities With
                                        {.QuantityType = qntyTyp,
                                         .QuntityForLinear = 2,
                                         .TargetMultiplier = tgtMul,
                                         .TypeOfExit = HKPositionalStrategyRule1.ExitType.CompoundingToMonthlyATR}

                                    filename = String.Format("CNC Wick Capital {0},QuantityType {1},TargetMul {2}",
                                                           If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital),
                                                           CType(.RuleEntityData, HKPositionalStrategyRule1.StrategyRuleEntities).QuantityType,
                                                           CType(.RuleEntityData, HKPositionalStrategyRule1.StrategyRuleEntities).TargetMultiplier)
                            End Select

                            .NumberOfTradeableStockPerDay = 10

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
#End Region

#Region "SMI HK Positional Strategy Rule"
            'Using backtestStrategy As New CNCEODGenericStrategy(canceller:=_canceller,
            '                                            exchangeStartTime:=TimeSpan.Parse("09:15:00"),
            '                                            exchangeEndTime:=TimeSpan.Parse("15:29:59"),
            '                                            tradeStartTime:=TimeSpan.Parse("09:15:00"),
            '                                            lastTradeEntryTime:=TimeSpan.Parse("15:29:59"),
            '                                            eodExitTime:=TimeSpan.Parse("15:29:59"),
            '                                            tickSize:=tick,
            '                                            marginMultiplier:=margin,
            '                                            timeframe:=1,
            '                                            heikenAshiCandle:=False,
            '                                            stockType:=stockType,
            '                                            databaseTable:=database,
            '                                            dataSource:=sourceData,
            '                                            initialCapital:=Decimal.MaxValue / 2,
            '                                            usableCapital:=Decimal.MaxValue / 2,
            '                                            minimumEarnedCapitalToWithdraw:=Decimal.MaxValue / 2,
            '                                            amountToBeWithdrawn:=50000)
            '    AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

            '    With backtestStrategy
            '        '.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Vijay CNC Instrument Details.csv")
            '        .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Investment Stock List.csv")

            '        .RuleNumber = GetComboBoxIndex_ThreadSafe(cmbRule)

            '        .RuleEntityData = New SMIHKPositionalStrategyRule.StrategyRuleEntities With
            '            {.QuantityType = SMIHKPositionalStrategyRule.TypeOfQuantity.Linear,
            '             .QuntityForLinear = 2}

            '        .NumberOfTradeableStockPerDay = 1

            '        .NumberOfTradesPerDay = Integer.MaxValue
            '        .NumberOfTradesPerStockPerDay = Integer.MaxValue

            '        .TickBasedStrategy = True
            '    End With
            '    Dim filename As String = String.Format("CNC Output Capital {3} {0}_{1}_{2}", Now.Hour, Now.Minute, Now.Second,
            '                                       If(backtestStrategy.UsableCapital = Decimal.MaxValue / 2, "∞", backtestStrategy.UsableCapital))
            '    Await backtestStrategy.TestStrategyAsync(startDate, endDate, filename).ConfigureAwait(False)
            'End Using
#End Region

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

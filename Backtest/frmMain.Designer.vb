<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.btnStart = New System.Windows.Forms.Button()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.grpbxDataSource = New System.Windows.Forms.GroupBox()
        Me.rdbLive = New System.Windows.Forms.RadioButton()
        Me.rdbDatabase = New System.Windows.Forms.RadioButton()
        Me.cmbRule = New System.Windows.Forms.ComboBox()
        Me.lblChooseRule = New System.Windows.Forms.Label()
        Me.lblStartDate = New System.Windows.Forms.Label()
        Me.lblEndDate = New System.Windows.Forms.Label()
        Me.dtpckrStartDate = New System.Windows.Forms.DateTimePicker()
        Me.dtpckrEndDate = New System.Windows.Forms.DateTimePicker()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.grpbxStrategyType = New System.Windows.Forms.GroupBox()
        Me.rdbCNC = New System.Windows.Forms.RadioButton()
        Me.rdbMIS = New System.Windows.Forms.RadioButton()
        Me.grpbxDataSource.SuspendLayout()
        Me.grpbxStrategyType.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStart.Location = New System.Drawing.Point(151, 142)
        Me.btnStart.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(181, 57)
        Me.btnStart.TabIndex = 0
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(7, 230)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(665, 50)
        Me.lblProgress.TabIndex = 1
        Me.lblProgress.Text = "Progress Status ....."
        '
        'grpbxDataSource
        '
        Me.grpbxDataSource.Controls.Add(Me.rdbLive)
        Me.grpbxDataSource.Controls.Add(Me.rdbDatabase)
        Me.grpbxDataSource.Location = New System.Drawing.Point(476, 9)
        Me.grpbxDataSource.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpbxDataSource.Name = "grpbxDataSource"
        Me.grpbxDataSource.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpbxDataSource.Size = New System.Drawing.Size(187, 57)
        Me.grpbxDataSource.TabIndex = 24
        Me.grpbxDataSource.TabStop = False
        Me.grpbxDataSource.Text = "Data Source"
        '
        'rdbLive
        '
        Me.rdbLive.AutoSize = True
        Me.rdbLive.Location = New System.Drawing.Point(115, 25)
        Me.rdbLive.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rdbLive.Name = "rdbLive"
        Me.rdbLive.Size = New System.Drawing.Size(55, 21)
        Me.rdbLive.TabIndex = 1
        Me.rdbLive.Text = "Live"
        Me.rdbLive.UseVisualStyleBackColor = True
        '
        'rdbDatabase
        '
        Me.rdbDatabase.AutoSize = True
        Me.rdbDatabase.Checked = True
        Me.rdbDatabase.Location = New System.Drawing.Point(7, 23)
        Me.rdbDatabase.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rdbDatabase.Name = "rdbDatabase"
        Me.rdbDatabase.Size = New System.Drawing.Size(90, 21)
        Me.rdbDatabase.TabIndex = 0
        Me.rdbDatabase.TabStop = True
        Me.rdbDatabase.Text = "Database"
        Me.rdbDatabase.UseVisualStyleBackColor = True
        '
        'cmbRule
        '
        Me.cmbRule.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbRule.FormattingEnabled = True
        Me.cmbRule.Items.AddRange(New Object() {"Smallest Candle Breakout", "High Volume Pin Bar", "Momentum Reversal v2", "High Volume Pin Bar v2", "Donchian Fractal Breakout", "SMI Fractal Breakout", "Day Long SMI (BANKNIFTY)", "Day Start SMI", "Gap Fractal Breakout", "Forward Momentum", "Vijay CNC", "TII Opposite Breakout", "Fixed Level Based", "Low Stoploss", "Multi Target", "Reversal", "Pinbar Breakout", "Low SL Pinbar", "Investment CNC", "Low Stoploss Wick", "Low Stoploss Candle", "Pair Trading", "Coin Flip At Resistance", "HeikenAshi CNC", "HeikenAshi CNC-1", "SMI HeikenAshi CNC", "HeikenAshi Hourly CNC-1", "ATR CNC"})
        Me.cmbRule.Location = New System.Drawing.Point(123, 30)
        Me.cmbRule.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cmbRule.Name = "cmbRule"
        Me.cmbRule.Size = New System.Drawing.Size(333, 26)
        Me.cmbRule.TabIndex = 22
        '
        'lblChooseRule
        '
        Me.lblChooseRule.AutoSize = True
        Me.lblChooseRule.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblChooseRule.Location = New System.Drawing.Point(12, 33)
        Me.lblChooseRule.Name = "lblChooseRule"
        Me.lblChooseRule.Size = New System.Drawing.Size(95, 18)
        Me.lblChooseRule.TabIndex = 23
        Me.lblChooseRule.Text = "Choose Rule"
        '
        'lblStartDate
        '
        Me.lblStartDate.AutoSize = True
        Me.lblStartDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStartDate.Location = New System.Drawing.Point(12, 89)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(74, 18)
        Me.lblStartDate.TabIndex = 25
        Me.lblStartDate.Text = "Start Date"
        '
        'lblEndDate
        '
        Me.lblEndDate.AutoSize = True
        Me.lblEndDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEndDate.Location = New System.Drawing.Point(241, 89)
        Me.lblEndDate.Name = "lblEndDate"
        Me.lblEndDate.Size = New System.Drawing.Size(69, 18)
        Me.lblEndDate.TabIndex = 26
        Me.lblEndDate.Text = "End Date"
        '
        'dtpckrStartDate
        '
        Me.dtpckrStartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpckrStartDate.Location = New System.Drawing.Point(96, 86)
        Me.dtpckrStartDate.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpckrStartDate.Name = "dtpckrStartDate"
        Me.dtpckrStartDate.Size = New System.Drawing.Size(137, 22)
        Me.dtpckrStartDate.TabIndex = 27
        '
        'dtpckrEndDate
        '
        Me.dtpckrEndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpckrEndDate.Location = New System.Drawing.Point(321, 86)
        Me.dtpckrEndDate.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpckrEndDate.Name = "dtpckrEndDate"
        Me.dtpckrEndDate.Size = New System.Drawing.Size(137, 22)
        Me.dtpckrEndDate.TabIndex = 28
        '
        'btnStop
        '
        Me.btnStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStop.Location = New System.Drawing.Point(340, 142)
        Me.btnStop.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(181, 57)
        Me.btnStop.TabIndex = 1
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'grpbxStrategyType
        '
        Me.grpbxStrategyType.Controls.Add(Me.rdbCNC)
        Me.grpbxStrategyType.Controls.Add(Me.rdbMIS)
        Me.grpbxStrategyType.Location = New System.Drawing.Point(476, 66)
        Me.grpbxStrategyType.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpbxStrategyType.Name = "grpbxStrategyType"
        Me.grpbxStrategyType.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpbxStrategyType.Size = New System.Drawing.Size(187, 57)
        Me.grpbxStrategyType.TabIndex = 29
        Me.grpbxStrategyType.TabStop = False
        Me.grpbxStrategyType.Text = "Strategy Type"
        '
        'rdbCNC
        '
        Me.rdbCNC.AutoSize = True
        Me.rdbCNC.Location = New System.Drawing.Point(115, 25)
        Me.rdbCNC.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rdbCNC.Name = "rdbCNC"
        Me.rdbCNC.Size = New System.Drawing.Size(57, 21)
        Me.rdbCNC.TabIndex = 1
        Me.rdbCNC.Text = "CNC"
        Me.rdbCNC.UseVisualStyleBackColor = True
        '
        'rdbMIS
        '
        Me.rdbMIS.AutoSize = True
        Me.rdbMIS.Checked = True
        Me.rdbMIS.Location = New System.Drawing.Point(7, 23)
        Me.rdbMIS.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rdbMIS.Name = "rdbMIS"
        Me.rdbMIS.Size = New System.Drawing.Size(52, 21)
        Me.rdbMIS.TabIndex = 0
        Me.rdbMIS.TabStop = True
        Me.rdbMIS.Text = "MIS"
        Me.rdbMIS.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(679, 293)
        Me.Controls.Add(Me.grpbxStrategyType)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.dtpckrEndDate)
        Me.Controls.Add(Me.dtpckrStartDate)
        Me.Controls.Add(Me.lblEndDate)
        Me.Controls.Add(Me.lblStartDate)
        Me.Controls.Add(Me.grpbxDataSource)
        Me.Controls.Add(Me.cmbRule)
        Me.Controls.Add(Me.lblChooseRule)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.btnStart)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Algo2Trade Backtest"
        Me.grpbxDataSource.ResumeLayout(False)
        Me.grpbxDataSource.PerformLayout()
        Me.grpbxStrategyType.ResumeLayout(False)
        Me.grpbxStrategyType.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnStart As Button
    Friend WithEvents lblProgress As Label
    Friend WithEvents grpbxDataSource As GroupBox
    Friend WithEvents rdbLive As RadioButton
    Friend WithEvents rdbDatabase As RadioButton
    Friend WithEvents cmbRule As ComboBox
    Friend WithEvents lblChooseRule As Label
    Friend WithEvents lblStartDate As Label
    Friend WithEvents lblEndDate As Label
    Friend WithEvents dtpckrStartDate As DateTimePicker
    Friend WithEvents dtpckrEndDate As DateTimePicker
    Friend WithEvents btnStop As Button
    Friend WithEvents grpbxStrategyType As GroupBox
    Friend WithEvents rdbCNC As RadioButton
    Friend WithEvents rdbMIS As RadioButton
End Class

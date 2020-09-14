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
        Me.rdbCNCEOD = New System.Windows.Forms.RadioButton()
        Me.rdbCNCCandle = New System.Windows.Forms.RadioButton()
        Me.rdbCNCTick = New System.Windows.Forms.RadioButton()
        Me.rdbMIS = New System.Windows.Forms.RadioButton()
        Me.grpBxDBConnection = New System.Windows.Forms.GroupBox()
        Me.rdbRemoteDBConnection = New System.Windows.Forms.RadioButton()
        Me.rdbLocalDBConnection = New System.Windows.Forms.RadioButton()
        Me.grpbxDataSource.SuspendLayout()
        Me.grpbxStrategyType.SuspendLayout()
        Me.grpBxDBConnection.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStart.Location = New System.Drawing.Point(486, 57)
        Me.btnStart.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(181, 57)
        Me.btnStart.TabIndex = 0
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(7, 210)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(665, 77)
        Me.lblProgress.TabIndex = 1
        Me.lblProgress.Text = "Progress Status ....."
        '
        'grpbxDataSource
        '
        Me.grpbxDataSource.Controls.Add(Me.rdbLive)
        Me.grpbxDataSource.Controls.Add(Me.rdbDatabase)
        Me.grpbxDataSource.Location = New System.Drawing.Point(15, 140)
        Me.grpbxDataSource.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpbxDataSource.Name = "grpbxDataSource"
        Me.grpbxDataSource.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpbxDataSource.Size = New System.Drawing.Size(190, 62)
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
        Me.cmbRule.Items.AddRange(New Object() {"Reversal HHLL Breakout", "Fractal Trend Line", "Market Plus Market Minus", "Highest Lowest Point", "Heikenashi Reverse Slab", "EMA Scalping", "Supertrend Cut Reversal", "BNF Martingale Strategy", "HigherLow LowerHigh Breakout", "Always in Trade Martingale Strategy", "Martingale Strategy", "Anchor Satellite HK Strategy", "Small Opening Range Breakout", "Loss Makeup Favourable Fractal Breakout", "HK Reverse Slab Martingale Strategy", "Low Price Option Buy Only Strategy", "Low Price Option OI Change Buy Only Strategy", "Lower Price Option Buy Only EOD Strategy", "Every Minute Top Gainer Losser HK Reversal Strategy", "Loss Makeup Rainbow Strategy", "Anchor Satellite Loss Makeup Strategy", "Anchor Satellite Loss Makeup HK Strategy", "Loss Makeup Neutral Slab Strategy", "Neutral Slab Martingale Strategy", "Anchor Satellite Loss Makeup HK Futures Strategy", "HK Reversal Single Trade Strategy", "Momentum Reversal Strategy", "Stochastic Divergence Strategy", "Pair Anchor Satellite Loss Makeup HK Strategy", "Two Third Strategy", "Small Range Breakout", "Highest Lowest Point Anchor Satellite Strategy", "EMA SMA Crossover Strategy", "Bollinger Close Strategy", "Outside VWAP Strategy", "Higher Timeframe Signal Martingale Strategy", "Squeeze Breakout Strategy", "Multi Trade Loss Makeup Strategy", "Momentum Reversal Modified Strategy", "HK Reverse Exit Strategy", "Both Direction Multi Trades HK Strategy", "Market Entry Strategy", "HK Reversal Loss Makeup Strategy", "Supertrend Cut Strategy", "Both Direction Multi Trades Strategy", "Day High Low Swing Trendline Strategy", "Fibonacci Backtest Strategy", "Swing At Day High Low Strategy", "Fibonacci Opening Range Breakout Strategy", "Previous Day HK Trend Strategy", "Previous Day HK Trend Bollinger Strategy", "EMA Attraction Strategy", "Buy Below Fractal Strategy", "HK Reversal Martingale Strategy", "HK Reversal Adaptive Martingale Strategy", "Bollinger Touch Strategy", "At The Money Option Buy Only Strategy", "Previous Day HK Trend Swing Strategy"})
        Me.cmbRule.Location = New System.Drawing.Point(123, 16)
        Me.cmbRule.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cmbRule.Name = "cmbRule"
        Me.cmbRule.Size = New System.Drawing.Size(544, 26)
        Me.cmbRule.TabIndex = 22
        '
        'lblChooseRule
        '
        Me.lblChooseRule.AutoSize = True
        Me.lblChooseRule.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblChooseRule.Location = New System.Drawing.Point(12, 19)
        Me.lblChooseRule.Name = "lblChooseRule"
        Me.lblChooseRule.Size = New System.Drawing.Size(95, 18)
        Me.lblChooseRule.TabIndex = 23
        Me.lblChooseRule.Text = "Choose Rule"
        '
        'lblStartDate
        '
        Me.lblStartDate.AutoSize = True
        Me.lblStartDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStartDate.Location = New System.Drawing.Point(12, 57)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(78, 18)
        Me.lblStartDate.TabIndex = 25
        Me.lblStartDate.Text = "Start Date:"
        '
        'lblEndDate
        '
        Me.lblEndDate.AutoSize = True
        Me.lblEndDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEndDate.Location = New System.Drawing.Point(241, 57)
        Me.lblEndDate.Name = "lblEndDate"
        Me.lblEndDate.Size = New System.Drawing.Size(73, 18)
        Me.lblEndDate.TabIndex = 26
        Me.lblEndDate.Text = "End Date:"
        '
        'dtpckrStartDate
        '
        Me.dtpckrStartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpckrStartDate.Location = New System.Drawing.Point(96, 54)
        Me.dtpckrStartDate.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpckrStartDate.Name = "dtpckrStartDate"
        Me.dtpckrStartDate.Size = New System.Drawing.Size(137, 22)
        Me.dtpckrStartDate.TabIndex = 27
        '
        'dtpckrEndDate
        '
        Me.dtpckrEndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpckrEndDate.Location = New System.Drawing.Point(321, 54)
        Me.dtpckrEndDate.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpckrEndDate.Name = "dtpckrEndDate"
        Me.dtpckrEndDate.Size = New System.Drawing.Size(137, 22)
        Me.dtpckrEndDate.TabIndex = 28
        '
        'btnStop
        '
        Me.btnStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStop.Location = New System.Drawing.Point(486, 129)
        Me.btnStop.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(181, 57)
        Me.btnStop.TabIndex = 1
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'grpbxStrategyType
        '
        Me.grpbxStrategyType.Controls.Add(Me.rdbCNCEOD)
        Me.grpbxStrategyType.Controls.Add(Me.rdbCNCCandle)
        Me.grpbxStrategyType.Controls.Add(Me.rdbCNCTick)
        Me.grpbxStrategyType.Controls.Add(Me.rdbMIS)
        Me.grpbxStrategyType.Location = New System.Drawing.Point(15, 83)
        Me.grpbxStrategyType.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpbxStrategyType.Name = "grpbxStrategyType"
        Me.grpbxStrategyType.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpbxStrategyType.Size = New System.Drawing.Size(443, 56)
        Me.grpbxStrategyType.TabIndex = 29
        Me.grpbxStrategyType.TabStop = False
        Me.grpbxStrategyType.Text = "Strategy Type"
        '
        'rdbCNCEOD
        '
        Me.rdbCNCEOD.AutoSize = True
        Me.rdbCNCEOD.Location = New System.Drawing.Point(335, 20)
        Me.rdbCNCEOD.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rdbCNCEOD.Name = "rdbCNCEOD"
        Me.rdbCNCEOD.Size = New System.Drawing.Size(91, 21)
        Me.rdbCNCEOD.TabIndex = 3
        Me.rdbCNCEOD.Text = "CNC EOD"
        Me.rdbCNCEOD.UseVisualStyleBackColor = True
        '
        'rdbCNCCandle
        '
        Me.rdbCNCCandle.AutoSize = True
        Me.rdbCNCCandle.Location = New System.Drawing.Point(211, 20)
        Me.rdbCNCCandle.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rdbCNCCandle.Name = "rdbCNCCandle"
        Me.rdbCNCCandle.Size = New System.Drawing.Size(105, 21)
        Me.rdbCNCCandle.TabIndex = 2
        Me.rdbCNCCandle.Text = "CNC Candle"
        Me.rdbCNCCandle.UseVisualStyleBackColor = True
        '
        'rdbCNCTick
        '
        Me.rdbCNCTick.AutoSize = True
        Me.rdbCNCTick.Location = New System.Drawing.Point(98, 20)
        Me.rdbCNCTick.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rdbCNCTick.Name = "rdbCNCTick"
        Me.rdbCNCTick.Size = New System.Drawing.Size(87, 21)
        Me.rdbCNCTick.TabIndex = 1
        Me.rdbCNCTick.Text = "CNC Tick"
        Me.rdbCNCTick.UseVisualStyleBackColor = True
        '
        'rdbMIS
        '
        Me.rdbMIS.AutoSize = True
        Me.rdbMIS.Checked = True
        Me.rdbMIS.Location = New System.Drawing.Point(7, 20)
        Me.rdbMIS.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rdbMIS.Name = "rdbMIS"
        Me.rdbMIS.Size = New System.Drawing.Size(52, 21)
        Me.rdbMIS.TabIndex = 0
        Me.rdbMIS.TabStop = True
        Me.rdbMIS.Text = "MIS"
        Me.rdbMIS.UseVisualStyleBackColor = True
        '
        'grpBxDBConnection
        '
        Me.grpBxDBConnection.Controls.Add(Me.rdbRemoteDBConnection)
        Me.grpBxDBConnection.Controls.Add(Me.rdbLocalDBConnection)
        Me.grpBxDBConnection.Location = New System.Drawing.Point(211, 140)
        Me.grpBxDBConnection.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpBxDBConnection.Name = "grpBxDBConnection"
        Me.grpBxDBConnection.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpBxDBConnection.Size = New System.Drawing.Size(247, 62)
        Me.grpBxDBConnection.TabIndex = 30
        Me.grpBxDBConnection.TabStop = False
        Me.grpBxDBConnection.Text = "Database Connection"
        Me.grpBxDBConnection.Visible = False
        '
        'rdbRemoteDBConnection
        '
        Me.rdbRemoteDBConnection.AutoSize = True
        Me.rdbRemoteDBConnection.Location = New System.Drawing.Point(141, 25)
        Me.rdbRemoteDBConnection.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rdbRemoteDBConnection.Name = "rdbRemoteDBConnection"
        Me.rdbRemoteDBConnection.Size = New System.Drawing.Size(78, 21)
        Me.rdbRemoteDBConnection.TabIndex = 1
        Me.rdbRemoteDBConnection.Text = "Remote"
        Me.rdbRemoteDBConnection.UseVisualStyleBackColor = True
        '
        'rdbLocalDBConnection
        '
        Me.rdbLocalDBConnection.AutoSize = True
        Me.rdbLocalDBConnection.Checked = True
        Me.rdbLocalDBConnection.Location = New System.Drawing.Point(17, 23)
        Me.rdbLocalDBConnection.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rdbLocalDBConnection.Name = "rdbLocalDBConnection"
        Me.rdbLocalDBConnection.Size = New System.Drawing.Size(63, 21)
        Me.rdbLocalDBConnection.TabIndex = 0
        Me.rdbLocalDBConnection.TabStop = True
        Me.rdbLocalDBConnection.Text = "Local"
        Me.rdbLocalDBConnection.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(679, 291)
        Me.Controls.Add(Me.grpBxDBConnection)
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
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Algo2Trade Backtest"
        Me.grpbxDataSource.ResumeLayout(False)
        Me.grpbxDataSource.PerformLayout()
        Me.grpbxStrategyType.ResumeLayout(False)
        Me.grpbxStrategyType.PerformLayout()
        Me.grpBxDBConnection.ResumeLayout(False)
        Me.grpBxDBConnection.PerformLayout()
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
    Friend WithEvents rdbCNCTick As RadioButton
    Friend WithEvents rdbMIS As RadioButton
    Friend WithEvents rdbCNCEOD As RadioButton
    Friend WithEvents rdbCNCCandle As RadioButton
    Friend WithEvents grpBxDBConnection As GroupBox
    Friend WithEvents rdbRemoteDBConnection As RadioButton
    Friend WithEvents rdbLocalDBConnection As RadioButton
End Class

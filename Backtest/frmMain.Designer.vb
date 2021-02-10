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
        Me.cmbRule = New System.Windows.Forms.ComboBox()
        Me.lblChooseRule = New System.Windows.Forms.Label()
        Me.lblStartDate = New System.Windows.Forms.Label()
        Me.lblEndDate = New System.Windows.Forms.Label()
        Me.dtpckrStartDate = New System.Windows.Forms.DateTimePicker()
        Me.dtpckrEndDate = New System.Windows.Forms.DateTimePicker()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStart.Location = New System.Drawing.Point(113, 115)
        Me.btnStart.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(136, 46)
        Me.btnStart.TabIndex = 0
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(5, 187)
        Me.lblProgress.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(499, 41)
        Me.lblProgress.TabIndex = 1
        Me.lblProgress.Text = "Progress Status ....."
        '
        'cmbRule
        '
        Me.cmbRule.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbRule.FormattingEnabled = True
        Me.cmbRule.Items.AddRange(New Object() {"Pivot Trend Option Buy Mode 3 Strategy", "HK Trend Option Buy Mode 3 Strategy", "HK MA Trend Option Buy Mode 3 Strategy", "Central Pivot Trend Option Buy Mode 3 Strategy"})
        Me.cmbRule.Location = New System.Drawing.Point(92, 24)
        Me.cmbRule.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmbRule.Name = "cmbRule"
        Me.cmbRule.Size = New System.Drawing.Size(409, 23)
        Me.cmbRule.TabIndex = 22
        '
        'lblChooseRule
        '
        Me.lblChooseRule.AutoSize = True
        Me.lblChooseRule.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblChooseRule.Location = New System.Drawing.Point(9, 27)
        Me.lblChooseRule.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblChooseRule.Name = "lblChooseRule"
        Me.lblChooseRule.Size = New System.Drawing.Size(78, 15)
        Me.lblChooseRule.TabIndex = 23
        Me.lblChooseRule.Text = "Choose Rule"
        '
        'lblStartDate
        '
        Me.lblStartDate.AutoSize = True
        Me.lblStartDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStartDate.Location = New System.Drawing.Point(9, 72)
        Me.lblStartDate.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(64, 15)
        Me.lblStartDate.TabIndex = 25
        Me.lblStartDate.Text = "Start Date:"
        '
        'lblEndDate
        '
        Me.lblEndDate.AutoSize = True
        Me.lblEndDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEndDate.Location = New System.Drawing.Point(269, 72)
        Me.lblEndDate.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblEndDate.Name = "lblEndDate"
        Me.lblEndDate.Size = New System.Drawing.Size(61, 15)
        Me.lblEndDate.TabIndex = 26
        Me.lblEndDate.Text = "End Date:"
        '
        'dtpckrStartDate
        '
        Me.dtpckrStartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpckrStartDate.Location = New System.Drawing.Point(79, 71)
        Me.dtpckrStartDate.Name = "dtpckrStartDate"
        Me.dtpckrStartDate.Size = New System.Drawing.Size(122, 20)
        Me.dtpckrStartDate.TabIndex = 27
        '
        'dtpckrEndDate
        '
        Me.dtpckrEndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpckrEndDate.Location = New System.Drawing.Point(335, 71)
        Me.dtpckrEndDate.Name = "dtpckrEndDate"
        Me.dtpckrEndDate.Size = New System.Drawing.Size(121, 20)
        Me.dtpckrEndDate.TabIndex = 28
        '
        'btnStop
        '
        Me.btnStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStop.Location = New System.Drawing.Point(255, 115)
        Me.btnStop.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(136, 46)
        Me.btnStop.TabIndex = 1
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(509, 238)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.dtpckrEndDate)
        Me.Controls.Add(Me.dtpckrStartDate)
        Me.Controls.Add(Me.lblEndDate)
        Me.Controls.Add(Me.lblStartDate)
        Me.Controls.Add(Me.cmbRule)
        Me.Controls.Add(Me.lblChooseRule)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.btnStart)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Algo2Trade Backtest"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnStart As Button
    Friend WithEvents lblProgress As Label
    Friend WithEvents cmbRule As ComboBox
    Friend WithEvents lblChooseRule As Label
    Friend WithEvents lblStartDate As Label
    Friend WithEvents lblEndDate As Label
    Friend WithEvents dtpckrStartDate As DateTimePicker
    Friend WithEvents dtpckrEndDate As DateTimePicker
    Friend WithEvents btnStop As Button
End Class

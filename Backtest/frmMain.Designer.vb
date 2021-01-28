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
        Me.btnStart.Location = New System.Drawing.Point(491, 12)
        Me.btnStart.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(181, 57)
        Me.btnStart.TabIndex = 0
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(7, 155)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(665, 73)
        Me.lblProgress.TabIndex = 1
        Me.lblProgress.Text = "Progress Status ....."
        '
        'cmbRule
        '
        Me.cmbRule.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbRule.FormattingEnabled = True
        Me.cmbRule.Items.AddRange(New Object() {"Below Portfolio Value CNC Strategy", "At Previous Nifty 50 Swing Low CNC Strategy"})
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
        Me.btnStop.Location = New System.Drawing.Point(491, 73)
        Me.btnStop.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(181, 57)
        Me.btnStop.TabIndex = 1
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(679, 231)
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
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
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

Public Class frmDisplay
    Private _dt As DataTable
    Public Sub New(ByVal dt As DataTable)
        InitializeComponent()
        _dt = dt
    End Sub

    Private Sub frmDisplay_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dgv.DataSource = _dt
        dgv.Refresh()
    End Sub
End Class
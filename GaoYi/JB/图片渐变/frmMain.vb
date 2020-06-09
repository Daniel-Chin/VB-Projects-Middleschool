Public Class frmMain
    Dim f As Form

    Private Sub Form1_Click(sender As Object, e As System.EventArgs) Handles Me.Click
        f = New frmMain
        f.Size = Me.Size
        f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        f.BackColor = Color.Red
        f.TopMost = True
        f.Opacity = 0
        f.ShowInTaskbar = False
        f.Show()
        f.Location = PointToScreen(Me.ClientRectangle.Location)
        f.Opacity = 1
        Me.BackgroundImage = My.Resources.Hydrangeas
        tmr.Enabled = True
    End Sub

    Private Sub frmMain_Move(sender As Object, e As System.EventArgs) Handles Me.Move
        If f IsNot Nothing Then
            f.Left = PointToScreen(Me.ClientRectangle.Location).X
            f.Top = PointToScreen(Me.ClientRectangle.Location).Y
        End If
    End Sub

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.BackgroundImage = My.Resources.Chrysanthemum
    End Sub

    Private Sub tmr_Tick(sender As System.Object, e As System.EventArgs) Handles tmr.Tick
        f.Opacity -= 0.02
        If f.Opacity <= 0 Then f.Dispose() : tmr.Enabled = False
    End Sub
End Class

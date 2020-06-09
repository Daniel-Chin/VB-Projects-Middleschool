Attribute VB_Name = "Module1"
Private Sub Main()
    If InStr(Command, "once") Then
        Form1.Show 1
    Else
        Form1.Show
    End If
End Sub

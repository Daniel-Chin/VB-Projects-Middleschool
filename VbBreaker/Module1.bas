Attribute VB_Name = "Module1"
Private Sub Main()
Open "d:\vbbreak" For Binary As #1
Put #1, 1, 1
Close #1
End Sub

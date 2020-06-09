Attribute VB_Name = "Only"
Private Sub Main()
    Open "D:\isjack" For Binary As #1
    Put #1, 1, 255
    Close #1
End Sub


Attribute VB_Name = "Dick"
Private Sub main()
Dim Wd As String
XH: Wd = InputBox("����ʲô�ʻ㣿", "Oxford learners dic", "")
If Wd <> "" Then
    Shell "explorer http://www.oxfordlearnersdictionaries.com/definition/english/" & Wd
    GoTo XH
End If
End Sub

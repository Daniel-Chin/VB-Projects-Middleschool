VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   4668
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   4668
   ScaleWidth      =   7740
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim cMd As String, Code As Byte, i As Variant
cMd = Mid(Command, 2, Len(Command) - 2)
Open cMd For Binary As #1
    Select Case MsgBox("目标文件为 " & cMd & " 。您要解密(Y)还是加密(N)？", vbQuestion + vbYesNoCancel, "")
        Case vbYes
            For i = 1 To LOF(1)
                Get #1, i, Code
                Code = IIf(Code = 255, 0, Code + 1)
                Put #1, i, Code
            Next i
            MsgBox "成功。", vbInformation + vbOKOnly, ""
            Close #1
            End
        Case vbNo
            For i = 1 To LOF(1)
                Get #1, i, Code
                Code = IIf(Code = 0, 255, Code - 1)
                Put #1, i, Code
            Next i
            MsgBox "成功。", vbInformation + vbOKOnly, ""
            Close #1
            End
    End Select
    Close #1
    End
End Sub

VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "生僻字符ASCii码查询"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3996
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   3996
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1452
      Left            =   732
      TabIndex        =   0
      Top             =   432
      Width           =   2532
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Text1) = 1 Then
        MsgBox QiuMa(Text1), vbOKOnly, "查询结果"
    Else
        MsgBox "请输入单个字符。", vbCritical, "错误"
    End If
End If
End Sub
Public Function QiuMa(Text As String)
Dim i As Long
    For i = -29999 To 29999
        If Chr(i) = Text Then
            QiuMa = i
            Exit For
        End If
    Next i
End Function

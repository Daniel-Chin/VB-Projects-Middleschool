VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5784
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9516
   LinkTopic       =   "Form1"
   ScaleHeight     =   5784
   ScaleWidth      =   9516
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   192
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cmd As String

Private Sub Form_Load()
If Len(Command) <= 4 Then
    cmd = InputBox("文件路径=", "", "")
Else
    cmd = Command
End If
Open cmd For Binary As #1
Dim k As Long, c As Byte
Show
For k = 1 To LOF(1)
    Get 1, k, c
    Text1 = Text1 & c & ","
Next k
Close #1
MsgBox "done."
End Sub

Private Sub Form_Resize()
With Text1
.Width = Form1.Width - 500
.Height = Form1.Height
End With
End Sub

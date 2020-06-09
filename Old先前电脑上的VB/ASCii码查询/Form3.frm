VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "¶ÎÂäASCii×ª»»"
   ClientHeight    =   5772
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8544
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5772
   ScaleWidth      =   8544
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text2 
      Height          =   2532
      Left            =   246
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3000
      Width           =   8052
   End
   Begin VB.TextBox Text1 
      Height          =   2292
      Left            =   246
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form3.frx":058A
      Top             =   240
      Width           =   8052
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Form2.Visible = True
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Form2.Visible = False
If KeyAscii = 13 Then
    Text2 = ""
    For i = 1 To Len(Text1)
        Text2 = Form2.QiuMa(Left(Right(Text1, i), 1)) & " " & Text2
    Next i
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1 = ""
    Text2 = Text2 & " "
    If InStr(Text2, " ") = 0 Then
        GoTo JieShu
    Else
        GoTo XunHuan
    End If
XunHuan:    a = InStr(Text2, " ")
    Text1 = Text1 & Chr(Val(Left(Text2, a - 1)))
    Text2 = Right(Text2, Len(Text2) - a)
    If InStr(Text2, " ") = 0 Then
    Else
        GoTo XunHuan
    End If
End If
JieShu:
End Sub


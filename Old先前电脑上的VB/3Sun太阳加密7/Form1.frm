VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "太阳加密3.2代"
   ClientHeight    =   6216
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6216
   ScaleWidth      =   10800
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4920
      Top             =   2880
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6012
      Left            =   114
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   102
      Width           =   10572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FilePath As String
Dim Saved As Boolean
Dim T As Integer
Private Sub Save()
Open FilePath For Binary As #2
Dim Le As Integer
Le = Len(Text1)
Dim i As Integer
For i = 1 To Le
    Dim Code_1 As Byte, Code_2 As Byte
    Dim Code As Integer
    Code = 524 - QiuMa(Right(Left(Text1, i), 1))
    Code_1 = Code \ 256
    Code_2 = Code Mod 256
    Put #2, i * 2 - 1, Code_1
    Put #2, i * 2, Code_2
Next i
FUCK: Get #2, Le + T, Code
If Code <> 0 Then
    T = T + 1
    Put #2, Le + T, 0
    GoTo FUCK
End If
Close #2
Saved = True
End Sub

Private Function QiuMa(Text As String) As Integer
Dim i As Integer
    For i = -29999 To 255
        If Chr(i) = Text Then
            QiuMa = i
            Exit For
        End If
    Next i
End Function

Private Sub Form_Unload(Cancel As Integer)
If Saved = False Then
    Save
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 19     'Ctrl+S
        If Saved = False Then
            Call Save
        End If
    Case 15     'Ctrl+O
        Call Timer1_Timer
    Case 1      'Ctrl+A
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)
    Case Else
        Saved = False
End Select
End Sub

Private Sub Timer1_Timer()
Dim CD As Byte
CD = Len(Command)
If CD > 3 Then
    FilePath = Right(Left(Command, CD - 1), CD - 2)
    GoTo DaKai
End If
Text1 = ""
FilePath = "F:\Sun太阳纪\3.2\Files\" & InputBox("Which file to open?", "Choose file", Month(Now) & "-" & Day(Now)) & ".Qnf"
DaKai: Open FilePath For Binary As #1
Dim i As Integer
Dim Code_1 As Byte, Code_2 As Byte, Code As Integer
For i = 1 To 29999 Step 2
    Get #1, i, Code_1
    Get #1, i + 1, Code_2
    Code = Code_1 * 256 + Code_2
    If Code = 0 Then
        Exit For
    End If
    Text1 = Text1 & Chr(524 - Code)
Next i
Close #1
Timer1.Interval = 0
End Sub

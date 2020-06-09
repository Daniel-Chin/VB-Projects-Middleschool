VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "太阳加密6代"
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
Dim NewFile As Boolean

Private Sub Save()
If NewFile = True Then
    FilePath = InputBox("Where to save?", "Save as", "F:\Sun太阳纪\6\" & Month(Now) & "-" & Day(Now) & ".Qnf")
End If
On Error GoTo CuoWu
Open FilePath For Binary As #2
Dim i As Integer
For i = 1 To Len(Text1)
    Dim Code_1 As Byte, Code_2 As Byte
    Dim Code As Integer
    Code = 524 - QiuMa(Right(Left(Text1, i), 1))
    Code_1 = Code \ 256
    Code_2 = Code Mod 256
    Put #2, i * 2 - 1, Code_1
    Put #2, i * 2, Code_2
Next i
Close #2
Saved = True
If 0 Then
CuoWu: MsgBox "路径不正确或文件名不合法。文件没有保存。", vbCritical, "Error"
End If
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
    Case 14     'Ctrl+N
        Call Save
        Shell App.Path & "\Sun.exe"
        End
    Case 1      'Ctrl+A
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)
    Case Else
        Saved = False
End Select
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 0
Text1 = ""
Dim CD As Byte
CD = Len(Command)
If CD > 3 Then
    FilePath = Right(Left(Command, CD - 1), CD - 2)
Else
    NewFile = True
    Exit Sub
End If
Open FilePath For Binary As #1
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
Saved = True
End Sub

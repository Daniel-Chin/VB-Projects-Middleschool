VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "太阳加密3.6代"
   ClientHeight    =   6216
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6216
   ScaleWidth      =   10800
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "解锁"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6252
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   10812
   End
   Begin VB.TextBox Tiao 
      BackColor       =   &H00C00000&
      Height          =   612
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Timer Saver 
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
   Begin VB.Timer Timer1 
      Left            =   10520
      Top             =   2880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileName As Integer
Public Copy As String
Dim P As Long
Dim Mh As Integer, Mw As Integer
Dim Saved As Boolean, Saving As Boolean, Oing As Boolean

Private Sub Command1_Click()
If InputBox("Password?", "Login", "") = "sh098-01B" Then
    Command1.Visible = False
Else
    MsgBox "Login failed.", vbOKOnly, "Error"
End If
End Sub

Private Sub Form_Load()
Form2.Visible = False
If Copy <> "" Then
    Text1 = Copy
    Copy = ""
    Exit Sub
End If
Dim RiQi As String
Retry: RiQi = InputBox("Which file to open?", "Choose file", Month(Now) & "/" & Day(Now))
Dim Yue As Byte, Ri As Byte
Yue = Val(Left(RiQi, InStr(RiQi, "/")))
Ri = Val(Right(RiQi, Len(RiQi) - InStr(RiQi, "/")))
If InStr(RiQi, "/") = 0 Or Yue <= 0 Or Yue > 12 Or Ri <= 0 Or Ri > 31 Then
    If MsgBox("Please input the file you want to open and press “OK”.", vbCritical + vbRetryCancel, "Error") = vbCancel Then
        End
    End If
    GoTo Retry
End If
FileName = Yue * 31 + Ri
Open App.Path & "/Files/" & FileName & ".dat" For Binary As #1
Dim Checker As Byte
Get #1, 1, Checker
If Checker <> 254 Then
    If MsgBox("No such file. Want to new one?", vbQuestion + vbOKCancel, "New") = vbOK Then
        Put #1, 1, 254
    Else
        Close #1
        GoTo Retry
    End If
End If
P = 1
Timer1.Interval = 1
Mw = Me.Width
Mh = Me.Height
Text1.Locked = True
Oing = True
End Sub

Private Sub Form_Resize()
With Me
    .Height = Mh
    .Width = Mw
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Saving = True Then
    MsgBox "Error! Save failed. Please retry.", vbCritical + vbOKOnly, "Error"
    Saver.Interval = 0
    GoTo Quxiao
End If
If Oing = True Then
    MsgBox "Error! Open failed. Please retry.", vbCritical + vbOKOnly, "Error"
    End
End If
If Saved = False Then
    Select Case MsgBox("File has been changed. Do you want to save it or not?", vbQuestion + vbYesNoCancel, "Not saves yet!")
        Case 6      'Yes
            Call save2
            End
        Case 7      'No
            End
        Case 2      'Cancel
Quxiao:            Form2.Timer1.Interval = 1
            Copy = Text1
    End Select
Else
    End
End If
End Sub

Private Sub Save()
If Saved = True Then
    Exit Sub
End If
Tiao.Visible = True
Tiao.Width = 0
Text1.Locked = True
Saver.Interval = 1
Saving = True
Open App.Path & "/Files/" & FileName & ".dat" For Binary As #2
End Sub

Private Sub save2()
Open App.Path & "/Files/" & FileName & ".dat" For Binary As #2
Rt: P = P + 1
If P - 1 > Len(Text1) Then
    Close #2
    Exit Sub
End If
Dim Code_1 As Byte, Code_2 As Byte
Dim Code As Integer
Code = 254 - QiuMa(Right(Left(Text1, P - 1), 1))
Code_1 = Code \ 256
Code_2 = Code Mod 256
Put #2, P * 2 - 2, Code_1
Put #2, P * 2 - 1, Code_2
GoTo Rt
End Sub

Private Sub Saver_Timer()
P = P + 1
Tiao.Width = P / Len(Text1) * Form1.Width
If P - 1 > Len(Text1) Then
    Saved = True
    Saving = False
    Saver.Interval = 0
    Text1.Locked = False
    Tiao.Visible = False
    Close #2
    P = 1
    Exit Sub
End If
Dim Code_1 As Byte, Code_2 As Byte
Dim Code As Integer
Code = 254 - QiuMa(Right(Left(Text1, P - 1), 1))
Code_1 = Code \ 256
Code_2 = Code Mod 256
Put #2, P * 2 - 2, Code_1
Put #2, P * 2 - 1, Code_2
End Sub

Private Function QiuMa(Text As String)
Dim i As Long
    For i = -29999 To 255
        If Chr(i) = Text Then
            QiuMa = i
            Exit For
        End If
    Next i
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Saving = True Then
    Exit Sub
End If
Select Case KeyAscii
    Case 19     'Ctrl+S
        If Saved = flase Then
            Call Save
        End If
    Case 12     'Ctrl+L
        Command1.Visible = True
    Case 1      'Ctrl+A
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)
    Case Else
        Saved = False
End Select
End Sub

Private Sub Timer1_Timer()
P = P + 1
Dim Code_1 As Byte, Code_2 As Byte
Get #1, P, Code_1
P = P + 1
Get #1, P, Code_2
If Code_1 + Code_2 = 0 Then
    MsgBox "File opened.", vbInformation + vbOKOnly, "Information"
    Timer1.Interval = 0
    P = 1
    Close #1
    Text1.SelStart = 0
    Oing = False
Else
    Text1 = Text1 & Chr(254 - Code_1 * 256 - Code_2)
    Text1.SelStart = Len(Text1)
End If
End Sub

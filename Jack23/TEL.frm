VERSION 5.00
Begin VB.Form TEL 
   Caption         =   "杰克电话簿"
   ClientHeight    =   4980
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10620
   Icon            =   "TEL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   10620
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Pad 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   25.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4000
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4000
   End
End
Attribute VB_Name = "TEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Const WakerInt As Integer = 36

Dim D(1 To 666, 0 To 2) As String
Dim tP As Integer
Const iNs As String = "键入姓名拼音首字母来检索。例：qnf"
Dim oRd As String

Private Sub Form_Load()
Pad = iNs
'加载语音系统等系统。
Load AI

'读取电话簿
Dim a As String, MH As Integer, k As Integer
Dim GB As Integer
Open App.Path & "\r\tel.txt" For Input As #64
Do Until EOF(64)
    GB = GB + 1
    If GB = 667 Then
        MsgBox "电话簿记录数超过上限", vbCritical + vbOKOnly, "Error"
        Unload Me
    End If
    Input #64, a
    For k = 0 To 1
        MH = InStr(a, ":")
        D(GB, k) = Left(a, MH - 1)
        a = Mid(a, MH + 1, 99)
    Next k
    D(GB, 2) = a
Loop
tP = GB
Close #64
End Sub

Private Sub Form_Resize()
With Pad
    .Width = Width
    .Height = Height
End With
End Sub

Private Sub Pad_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Then
    KeyAscii = KeyAscii + 32    '大写转小写
End If
Select Case KeyAscii
    Case 27 'ESC
        Unload Me
    Case 5 'Ctrl+E
        Shell "explorer " & App.Path & "\r\tel.txt"
        Unload Me
    Case 97 To 122 '小写字母
        oRd = Right(oRd & Chr(KeyAscii), 6)
        XoRd
    Case 13 'Enter
        Pad = ""
        XoRd
    Case 48 To 57   '数字
        oRd = Right(oRd & Chr(KeyAscii), 6)
        XoRd
End Select
End Sub

Private Sub XoRd()
Dim OldPad As String
OldPad = Pad
Pad = iNs
Dim LineNo As Integer
LineNo = Val(Right(oRd, 1))
Dim k As Integer
For k = 1 To tP
    If InStr(oRd, D(k, 0)) <> 0 Then
        LineNo = LineNo - 1
        If LineNo <= 0 Then
            Pad = Pad & Chr(13) & Chr(10) & D(k, 1) & ":" & D(k, 2)
            If LineNo = 0 Then
                Exit For
            End If
        End If
    End If
Next k
If OldPad <> Pad Then
    AI.Mute
    AI.Speak Mid(Pad, Len(iNs) + 1, 9999)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
DoEvents
Unload AI
End Sub

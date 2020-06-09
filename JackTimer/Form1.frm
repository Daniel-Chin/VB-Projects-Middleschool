VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   Caption         =   "x"
   ClientHeight    =   6288
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10296
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6288
   ScaleWidth      =   10296
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Timer Main 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5520
      Top             =   3240
   End
   Begin VB.TextBox T 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2892
      Index           =   0
      Left            =   960
      MaxLength       =   2
      MousePointer    =   7  'Size N S
      TabIndex        =   0
      Text            =   "00"
      Top             =   1080
      Width           =   1812
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim k
Dim Off As Boolean
Dim Drag As Integer
Dim dS As Long, tS As Integer

Private Sub Form_DblClick()
If Not Main Then
    Me.BackColor = &HCC6600
    Main = True
    T(0).Locked = True
    T(1).Locked = True
    T(2).Locked = True
End If
End Sub

Private Sub Form_Load()
Drag = 99
Show
Load T(1)
Load T(2)
For k = 0 To 2
    With T(k)
        .FontSize = 166
        .Top = Me.Height / 6
        .Height = Me.Height / 2
        .Width = Me.Width / 4
        .Left = (1 / 16 + k * (1 / 16 + 1 / 4)) * Me.Width
        .Visible = True
    End With
Next k
T(1) = 15
T(0).SetFocus
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Main Then
    If Button <> 1 Then
        If MsgBox("定时结束之后关机？", vbYesNo + vbQuestion) = vbYes Then
            Off = True
            Me.BackColor = vbRed
            Main = True
            T(0).Locked = True
            T(1).Locked = True
            T(2).Locked = True
        End If
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Main Then
    Shell App.Path & "\" & App.EXEName, vbMaximizedFocus
End If
End Sub

Private Sub Main_Timer()
If T(2) = 0 Then
    If T(1) = 0 Then
        If T(0) = 0 Then
            If Off Then
Hehe:
                Shell "cmd /c shutdown -s -t 10"
                MsgBox "时间到！即将关机。按确定来取消关机。", vbOKOnly + vbExclamation
                Shell "cmd /c shutdown -a"
                If MsgBox("关机已取消。按“否”重新关机。", vbYesNo + vbQuestion) = vbNo Then
                    GoTo Hehe
                Else
                    End
                End If
            Else
                Alarm.Show 1
                Unload Me
                Exit Sub
            End If
        Else
            T(0) = T(0) - 1
        End If
    Else
        T(1) = T(1) - 1
    End If
    T(2) = 59
Else
    T(2) = T(2) - 1
End If
End Sub

Private Sub T_GotFocus(Index As Integer)
With T(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub T_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Form_DblClick
End If
End Sub

Private Sub T_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Drag = Index
dS = Y
tS = T(Index)
End Sub

Private Sub T_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Drag <> 99 And Not Main Then
    T(Index) = Abss(tS + (dS - Y) / 166)
End If
End Sub

Private Sub T_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Drag = 99
End Sub

Private Function Abss(A As Integer)
If A < 0 Then
    Abss = 0
Else
    Abss = A
End If
End Function

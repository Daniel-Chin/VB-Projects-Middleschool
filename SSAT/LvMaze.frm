VERSION 5.00
Begin VB.Form LvMaze 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "迷宫"
   ClientHeight    =   7032
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10836
   DrawWidth       =   3
   Icon            =   "LvMaze.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7032
   ScaleWidth      =   10836
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Main 
      Left            =   7200
      Top             =   3120
   End
   Begin VB.TextBox Focuser 
      Height          =   372
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   -666
      Width           =   972
   End
   Begin VB.Timer Killer 
      Interval        =   36
      Left            =   6840
      Top             =   2520
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "PLAY"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   4440
      MouseIcon       =   "LvMaze.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   2760
      Width           =   1212
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   252
      Index           =   22
      Left            =   8640
      TabIndex        =   24
      Top             =   600
      Width           =   1212
   End
   Begin VB.Label Fuck 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "WIN"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   9120
      TabIndex        =   23
      Top             =   120
      Width           =   612
   End
   Begin VB.Label Trap 
      BackColor       =   &H0000FFFF&
      Height          =   852
      Left            =   3600
      TabIndex        =   22
      Top             =   720
      Width           =   372
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   1092
      Index           =   21
      Left            =   2280
      TabIndex        =   21
      Top             =   720
      Width           =   852
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   20
      Left            =   720
      TabIndex        =   20
      Top             =   2160
      Width           =   1332
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   19
      Left            =   1920
      TabIndex        =   19
      Top             =   3000
      Width           =   1332
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   18
      Left            =   1080
      TabIndex        =   18
      Top             =   3480
      Width           =   1212
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   17
      Left            =   1560
      TabIndex        =   17
      Top             =   4320
      Width           =   1692
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   16
      Left            =   1200
      TabIndex        =   16
      Top             =   5040
      Width           =   1332
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   15
      Left            =   1560
      TabIndex        =   15
      Top             =   5640
      Width           =   1332
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   3492
      Index           =   14
      Left            =   9840
      TabIndex        =   14
      Top             =   0
      Width           =   972
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   732
      Index           =   13
      Left            =   960
      TabIndex        =   13
      Top             =   0
      Width           =   8052
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   6252
      Index           =   12
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   732
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   10692
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   3132
      Index           =   10
      Left            =   10320
      TabIndex        =   10
      Top             =   3360
      Width           =   732
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   9
      Left            =   9000
      TabIndex        =   9
      Top             =   3360
      Width           =   1452
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   1452
      Index           =   8
      Left            =   9960
      TabIndex        =   8
      Top             =   4320
      Width           =   132
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   252
      Index           =   7
      Left            =   3240
      TabIndex        =   7
      Top             =   5760
      Width           =   6852
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   492
      Index           =   6
      Left            =   4080
      TabIndex        =   6
      Top             =   4560
      Width           =   852
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   252
      Index           =   5
      Left            =   3960
      TabIndex        =   5
      Top             =   5040
      Width           =   5772
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   1812
      Index           =   4
      Left            =   6480
      TabIndex        =   4
      Top             =   3000
      Width           =   252
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   3372
      Index           =   3
      Left            =   9120
      TabIndex        =   3
      Top             =   2160
      Width           =   372
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   732
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   1560
      Width           =   6252
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   3732
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   2160
      Width           =   492
   End
   Begin VB.Label W 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   3840
      Width           =   4572
   End
End
Attribute VB_Name = "LvMaze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cx As Long, cy As Long
Dim bR As Byte, bb As Byte, BG As Byte
Dim Gaming As Boolean
Dim Tms As Integer
Dim pW As Boolean, pA As Boolean, pS As Boolean, pD As Boolean
Dim Death As Integer, aLf As Boolean

Private Sub Form_Unload(Cancel As Integer)
Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
End
End Sub

Private Sub Reset()
Cls
With W(21)
    .Top = 720
    .Left = 2280
    .Visible = False
End With
Trap.Visible = True
bb = 255
BG = 255
bR = 0
Gaming = False
Bt.Visible = True
With W(22)
    .Visible = False
    .Height = 252
End With
Killer.Enabled = False
Tms = 0
Focuser.SetFocus
End Sub

Private Sub Bt_Click()
Bt.Visible = False
Gaming = True
pW = False
pA = False
pS = False
pD = False
End Sub

Private Sub Bt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cx = X + Bt.Left
cy = Y + Bt.Top
End Sub

Private Sub Focuser_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Gaming Then
    Select Case KeyCode
        Case 38
            pW = True
        Case 37
            pA = True
        Case 40
            pS = True
        Case 39
            pD = True
    End Select
    Main_Timer
End If
End Sub

Private Sub Focuser_KeyUp(KeyCode As Integer, Shift As Integer)
If Not Gaming Then
    Select Case KeyCode
        Case 38
            pW = False
        Case 37
            pA = False
        Case 40
            pS = False
        Case 39
            pD = False
    End Select
End If
End Sub

Private Sub Form_Load()
aLf = False
Death = 0
Show
Reset
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Gaming Then
    Select Case 255
        Case bb
            If bR = 255 Then GoTo Gbr
            bR = bR + 1
            BG = BG - 1
        Case BG
            bb = bb + 1
            bR = bR - 1
        Case bR
Gbr:
            BG = BG + 1
            bb = bb - 1
    End Select
    Line (cx, cy)-(X, Y), RGB(bR, BG, bb)
    If Sqr((cx - X) ^ 2 + (cy - Y) ^ 2) >= 366 Then
        MsgBox "你刚才要么移动鼠标过快，要么使用了一些其他邪恶方法试图穿过墙壁。作弊是不好的行为哦！罚你重新开始。", vbExclamation + vbOKOnly, "迷宫"
        Reset
    End If
    cx = X
    cy = Y
End If
End Sub

Private Sub Fuck_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Gaming Then
    X = X + Fuck.Left
    Y = Y + Fuck.Top
    If Sqr((cx - X) ^ 2 + (cy - Y) ^ 2) >= 366 Then
        MsgBox "你刚才要么移动鼠标过快，要么使用了一些其他邪恶方法试图穿过墙壁。作弊是不好的行为哦！罚你重新开始。", vbExclamation + vbOKOnly, "迷宫"
        Reset
    Else
        Win
    End If
End If
End Sub

Private Sub Killer_Timer()
Tms = Tms + 1
Select Case Tms
    Case 1 To 28
        W(21).Left = W(21).Left + 16
    Case 29 To 126
        W(21).Left = W(21).Left + 66
    Case 136 To 666
        W(22).Height = W(22).Height + 16
End Select
Dim k As Integer
For k = 0 To 22
    W_MouseMove k, 0, 0, cx - W(k).Left, cy - W(k).Top
Next k
End Sub

Private Sub Main_Timer()
Bt.Top = Bt.Top + (pW - pS) * 26
Bt.Left = Bt.Left + (pA - pD) * 26
If (pW Or pS Or pA Or pD) And Not Gaming Then
    Main.Interval = 18
Else
    Main.Interval = 0
End If
End Sub

Private Sub Trap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Gaming Then
    aLf = True
    Trap.Visible = False
    W(21).Visible = True
    W(22).Visible = True
    Killer.Enabled = True
End If
End Sub

Private Sub W_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Gaming Then
    With W(Index)
        If X >= 0 And X <= .Width And Y >= 0 And Y <= .Height Then
            .BackColor = vbRed
            DoEvents
            Si
            .BackColor = 0
        End If
    End With
End If
End Sub

Private Sub Si()
MsgBox "你死了。没错，在这个游戏里，你不可以碰墙！", vbExclamation + vbOKOnly, "迷宫"
Death = Death + 1
If Death >= 10 And aLf Then
    MsgBox "你已经死了" & Death & "次了，还没有找到这一关的答案。按OK来查看提示。", vbOKOnly + vbInformation, "提示"
    MsgBox "这一关的关键，在于使用方向键操控。", vbInformation + vbOKOnly, "提示"
End If
Reset
End Sub

Private Sub Win()
SaveSetting "iGlope", "woqu", "j2", 1
MsgBox "经过" & Death & "次的尝试，你终于打通了奖励关卡“迷宫”。按OK回到主菜单。", vbInformation + vbOKOnly, "恭喜！"
Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
End
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MOD"
   ClientHeight    =   6756
   ClientLeft      =   36
   ClientTop       =   684
   ClientWidth     =   11556
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   42
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6756
   ScaleWidth      =   11556
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Restart 
      BackColor       =   &H00000000&
      Caption         =   "Restart"
      Height          =   840
      Left            =   3792
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   3972
   End
   Begin How2BecomeACalculater.Jdt Jdt 
      Height          =   612
      Left            =   1272
      Top             =   5640
      Width           =   9012
      _ExtentX        =   15896
      _ExtentY        =   1080
   End
   Begin VB.Timer Timer 
      Interval        =   36
      Left            =   5280
      Top             =   3240
   End
   Begin VB.TextBox Ans 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   3312
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   4200
      Width           =   4932
   End
   Begin VB.Label HS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HS"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Pad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Pad"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   5424
      TabIndex        =   5
      Top             =   1000
      Width           =   720
   End
   Begin VB.Label Board 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score"
      ForeColor       =   &H00FFFFFF&
      Height          =   888
      Left            =   4704
      TabIndex        =   4
      Top             =   0
      Width           =   2160
   End
   Begin VB.Label Midd 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mod"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   66
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   4776
      TabIndex        =   2
      Top             =   1680
      Width           =   2004
   End
   Begin VB.Label Num2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num2"
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   7098
      TabIndex        =   1
      Top             =   1920
      Width           =   1680
   End
   Begin VB.Label Num1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num1"
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   2778
      TabIndex        =   0
      Top             =   1920
      Width           =   1680
   End
   Begin VB.Menu Opt 
      Caption         =   "选项"
      Begin VB.Menu ResetHS 
         Caption         =   "重置最高分"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim Time As Single
Dim Score As Long
Dim NowScore As Long
Dim Die As Integer
Dim HighScore As Long
Dim MaxTime As Integer

Const MyName As String = "How 2 Become A Calculater"

Private Sub Ans_KeyPress(KeyAscii As Integer)
With Ans
    Select Case KeyAscii
        Case 1  'Ctrl+A
            .SelStart = 0
            .SelLength = Len(.Text)
        Case 13 'ENTER
            ENTER
        Case 27 'ESC
            .Text = ""
    End Select
End With
End Sub

Private Sub ENTER()
If Val(Ans) = Val(Num1) Mod Val(Num2) Then
    If Time > 5 Then
        Time = 5
    End If
    NowScore = NowScore + Int(Time) + 1
    If NowScore > Score Then
        PadLet "+" & (NowScore - Score)
        Score = NowScore
    End If
    Die = 0
    NXT
Else
    Beep
    Ans = ""
    Time = Time / 2
End If
End Sub

Private Sub NXT()
MaxTime = 7
Time = MaxTime - Die
Debug.Assert Time <> 0
If Time <= 0 Then
    Si
Else
    If Die = 0 Then
        Me.BackColor = 0
        Jdt.RenewBackColor 0, 0, 0
        Me.Caption = MyName
    Else
        Me.BackColor = Int(Die / MaxTime * 255)
        Jdt.RenewBackColor Me.BackColor, 0, 0
        Me.Caption = "累计失误次数：" & Die
    End If
    Ans = ""
    Dim Standard As Long
    Standard = (NowScore / 19) ^ 2.6
    If Rnd < 0.19 Then
        Num1 = Int(Rnd * Standard) + 3
    Else
        Num1 = Standard - Int(Rnd * Standard * 0.19) + 3
    End If
    Standard = Num1 * 0.1
    Num2 = Int(Standard * (1 + Rnd * 5))
    Num2 = Val(Num2) + Int(Rnd * 4) - 1
    If Val(Num2) > Val(Num1) Then
        Num2 = Num1
    Else
        If Val(Num2) <= 1 Then
            Num2 = 2
        End If
    End If
    Board = Score
End If
End Sub

Private Sub Form_Load()
Randomize
With Jdt
    .RenewBackColor 0, 0, 0
    .RenewBars 38
    .RenewEmptyColor 255, 255, 0
    .RenewFilledColor 0, 0, 255
    .RenewProgress 100
End With
NXT
Pad = ""
HighScore = GetSetting("iGlope", "MOD", "highscore", "0")
HS = "历史最高分：" & HighScore
End Sub

Private Sub PadLet(A As String)
With Pad
    If .BackColor = 0 Then
        .ForeColor = 0
        .BackColor = vbWhite
    Else
        .ForeColor = vbWhite
        .BackColor = 0
    End If
End With
Pad = A
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Score > HighScore Then
    SaveSetting "iGlope", "MOD", "highscore", Score
End If
End Sub

Private Sub ResetHS_Click()
If MsgBox("重置历史最高分吗？", vbQuestion + vbYesNo) = vbYes Then
    SaveSetting "iGlope", "MOD", "highscore", "0"
    MsgBox "重置完毕。即将重启游戏。", vbInformation + vbOKOnly
    Reboot
End If
Time = 0
End Sub

Private Sub Restart_Click()
Reboot
End Sub

Private Sub Timer_Timer()
Time = Time - Timer.Interval / 1000
If Time <= 0 Then
    NowScore = NowScore - MaxTime / 2
    If NowScore < 0 Then
        NowScore = 0
    End If
    With Ans
        .Text = Val(Num1) Mod Val(Num2)
        .BackColor = Me.BackColor
        .Alignment = 2
        .Enabled = False
    End With
    Beep
    DoEvents
    Sleep 1000
    DoEvents
    With Ans
        .Enabled = True
        .Alignment = 0
        .BackColor = vbWhite
    End With
    Die = Die + 1
    PadLet "时间减半！"
    NXT
End If
Jdt.RenewProgress Time / MaxTime * 100
End Sub

Private Sub Si()
Timer = False
Me.BackColor = 0
Me.Caption = MyName
Invisiblize Num1
Invisiblize Num2
Invisiblize Midd
Invisiblize Jdt
Invisiblize Ans
Invisiblize HS
Restart.Visible = True
Board = "分数：" & Score
Board.FontSize = 95
With Pad
    .Caption = "历史最高分：" & HighScore
    .FontSize = 66
    .Top = Board.Height + (Me.ScaleHeight - Board.Height - Pad.Height) * 0.6
    .ForeColor = vbWhite
    .BackColor = 0
End With
If Score > HighScore Then
    SaveSetting "iGlope", "MOD", "highscore", Score
    MsgBox "新纪录！", vbInformation + vbOKOnly
    Me.Caption = "新纪录：" & Score
End If
End Sub

Private Sub Invisiblize(Thing As Control)
Thing.Visible = False
End Sub

Private Sub Reboot()
Shell App.Path & "\" & App.EXEName, vbNormalFocus
Unload Me
End
End Sub

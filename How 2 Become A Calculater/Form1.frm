VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MOD"
   ClientHeight    =   6756
   ClientLeft      =   36
   ClientTop       =   684
   ClientWidth     =   13884
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
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6756
   ScaleWidth      =   13884
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Restart 
      BackColor       =   &H00000000&
      Caption         =   "Restart"
      Height          =   840
      Left            =   6528
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   3972
   End
   Begin How2BecomeACalculater.Jdt Jdt 
      Height          =   612
      Left            =   4008
      Top             =   5640
      Width           =   9012
      _ExtentX        =   15896
      _ExtentY        =   1080
   End
   Begin VB.Timer Timer 
      Interval        =   36
      Left            =   7440
      Top             =   3480
   End
   Begin VB.TextBox Ans 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   6048
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0CCA
      Top             =   4200
      Width           =   4932
   End
   Begin How2BecomeACalculater.Hexa Hexa1 
      Height          =   5052
      Left            =   480
      TabIndex        =   9
      Top             =   480
      Width           =   4212
      _ExtentX        =   7430
      _ExtentY        =   8911
   End
   Begin VB.Label Num1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num1"
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   5520
      TabIndex        =   0
      Top             =   1920
      Width           =   1680
   End
   Begin WMPLibCtl.WindowsMediaPlayer Music 
      Height          =   2892
      Left            =   5520
      TabIndex        =   10
      Top             =   -3000
      Visible         =   0   'False
      Width           =   2892
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5101
      _cy             =   5101
   End
   Begin VB.Label STG 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stage"
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
      TabIndex        =   8
      Top             =   0
      Width           =   1200
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
      Left            =   6582
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
      Left            =   5862
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
      Left            =   7512
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
      Left            =   9840
      TabIndex        =   1
      Top             =   1920
      Width           =   1680
   End
   Begin VB.Menu Opt 
      Caption         =   "选项"
      Begin VB.Menu MusicDown 
         Caption         =   "背景音乐音量-"
         Shortcut        =   {F11}
      End
      Begin VB.Menu MusicUp 
         Caption         =   "背景音乐音量+"
         Shortcut        =   {F12}
      End
      Begin VB.Menu ToggleFx 
         Caption         =   "关闭/打开音效"
      End
      Begin VB.Menu ResetHS 
         Caption         =   "重置最高分"
      End
      Begin VB.Menu Shuai 
         Caption         =   "开发者长得真帅！"
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
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function sndPlaySoundStop Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As Long, ByVal uFlags As Long) As Long

Dim Inf As Boolean
Dim Ending As Boolean
Dim EndTms As Single
Dim Fx As Boolean
Dim Vol As Integer
Dim Time As Single
Dim Score As Long
Dim NowScore As Long
Dim Die As Integer
Dim HighScore As Long
Dim MaxTime As Integer
Dim StageMap()
Dim Stage As Integer
Dim Silence As Byte

Const MyName As String = "How 2 Become A Calculater"

Private Sub Ans_KeyPress(KeyAscii As Integer)
With Ans
    If Not .Locked Then
        Select Case KeyAscii
            Case 1  'Ctrl+A
                .SelStart = 0
                .SelLength = Len(.Text)
            Case 13 'ENTER
                ENTER
            Case 27 'ESC
                .Text = ""
            Case 10 '!!!!!!!
                .Text = Val(Num1) Mod Val(Num2)
        End Select
    End If
End With
End Sub

Private Sub ENTER()
sndPlaySoundStop 0, 0
If Val(Ans) = Val(Num1) Mod Val(Num2) Then
    If Time > 5 Then
        Time = 5
    End If
    NowScore = NowScore + Int(Time) + 1
    WavPlay "c" & Int(Rnd() * 6)
    Dim AdvS As Boolean
    If NowScore > Score Then
        PadLet "+" & (NowScore - Score)
        Score = NowScore
        If Not Inf Then
            If Score >= StageMap(Stage - 1) Then
                If Stage = 13 Then
                    'Ending
                    If HighScore >= StageMap(12) Then
                        'Inf
                        Inf = True
                        STG = "-=|无尽挑战|=-"
                        STG.FontUnderline = True
                        Hexa1.SetMode 14
                    Else
                        Ans.Locked = True
                        Invisiblize STG
                        Invisiblize HS
                        Invisiblize Pad
                        Ending = True
                        sndPlaySoundStop 0, 0
                    End If
                Else
                    Stage = Stage + 1
                    STG = "STAGE " & Stage & " " & Chr(10) & "Next Stage:Score " & StageMap(Stage - 1) & " "
                    Hexa1.SetMode Stage
                    WavPlay "sc"
                    AdvS = True
                    Music.URL = App.Path & "\sound\m" & ((Stage + 1) Mod 3) + 1 & ".mp3"
                End If
            End If
        End If
    Else
        PadLet ""
    End If
    Me.Caption = "目前难度系数：" & NowScore
    Die = 0
    With Ans
        .Text = (Val(Num1) Mod Val(Num2)) & " "
        .BackColor = Me.BackColor
        .ForeColor = &HFFFF00
        .Alignment = 2
        .Locked = True
        .SelLength = 0
        .SelStart = Len(.Text)
    End With
    DoEvents
    If AdvS Then
        AdvS = False
        Sleep 1000
    End If
    Sleep 500
    DoEvents
    With Ans
        .Locked = False
        .Alignment = 0
        .BackColor = vbWhite
        .ForeColor = 0
        .SetFocus
    End With
    NXT
Else
    Beep
    Ans = ""
    Time = Time / 2
    PadLet "时间减半！"
    WavPlay Str(Round(Time))
End If
End Sub

Private Sub NXT()
MaxTime = 7
Time = MaxTime - Die
If Time <= 0 Then
    Si
Else
    WavPlay Str(Time)
    If Die = 0 Then
        Me.BackColor = 0
        Jdt.RenewBackColor 0, 0, 0
        Hexa1.SetBackColor 0
    Else
        Me.BackColor = Int(Die / MaxTime * 255)
        Jdt.RenewBackColor Me.BackColor, 0, 0
        Hexa1.SetBackColor Me.BackColor
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
    Num2 = Int(Standard * Rnd * 5)
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
Me.Caption = MyName
Fx = GetSetting("iGlope", "MOD", "fx", True)
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
StageMap = Array(25, 50, 70, 90, 115, 150, 170, 200, 225, 250, 300, 325, 360)
Stage = 1
STG.Left = Me.ScaleWidth - STG.Width
STG = "STAGE 1" & " " & Chr(10) & "Next Stage:Score " & StageMap(0) & " "
With Music
    Vol = Val(GetSetting("iGlope", "MOD", "volume", "10"))
    .settings.volume = Round((101 ^ 0.1) ^ Vol) - 1
    .URL = App.Path & "\sound\m" & Int(Rnd * 3 + 1) & ".mp3"
End With
If Vol = 0 Then
    MusicDown = False
End If
If Vol = 10 Then
    MusicUp = False
End If
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
sndPlaySoundStop 0, 0
If Score > HighScore Then
    SaveSetting "iGlope", "MOD", "highscore", Score
End If
If Ans.Locked Then
    End
End If
End Sub

Private Sub Hexa1_GotFocus()
If Ans.Visible And Not Ending Then
    Ans.SetFocus
End If
End Sub

Private Sub MusicUp_Click()
Vol = Vol + 1
MusicDown = True
If Vol = 10 Then
    MusicUp.Enabled = False
End If
Music.settings.volume = Round((101 ^ 0.1) ^ Vol) - 1
SaveSetting "iGlope", "MOD", "volume", Vol
End Sub

Private Sub MusicDown_Click()
Vol = Vol - 1
MusicUp = True
If Vol = 0 Then
    MusicDown.Enabled = False
End If
Music.settings.volume = Round((101 ^ 0.1) ^ Vol) - 1
SaveSetting "iGlope", "MOD", "volume", Vol
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

Private Sub Shuai_Click()
With Shuai
    .Caption = "谢谢，你也很帅！"
    .Enabled = False
End With
End Sub

Private Sub Timer_Timer()
If Not Ans.Locked Then
    If Ending Then
        If Int(EndTms) < 18 Then
            Jdt.RenewProgress Rnd * 100
            Num1 = Val(Num1) + Int(Rnd * 119) + 1
            Num2 = Int(Val(Num1) * Rnd) + 1
            Ans = Val(Num1) Mod Val(Num2)
            Board = Val(Board) + 1
        End If
        Select Case Int(EndTms)
            Case 0
                Fx = True
                Vol = 9
                MusicUp_Click
                sndPlaySoundStop 0, 0
                Hexa1.SetMode 14
                Music.URL = App.Path & "\sound\dalek.mp3"
                EndTms = EndTms + 1
                Timer.Interval = 400
            Case 2
                Timer.Interval = 200
            Case 3 To 5
                Timer.Interval = Timer.Interval * 0.8
                If Timer.Interval < 40 Then
                    Timer.Interval = 40
                End If
            Case 6
                Hexa1.SetMode 15
                EndTms = EndTms + 1
            Case 7 To 11
                Hexa1.Width = Hexa1.Width - Hexa1.Height * Timer.Interval / 22000
            Case 12 To 14
                Num1 = Int(Val(Num1) * 1.1)
            Case 15
                Music.Controls.stop
                EndTms = EndTms + 1
                sndPlaySoundStop 0, 0
                WavPlay "epic"
                Timer.Interval = 40
                Board.Visible = False
                Invisiblize Jdt
                Hexa1.SetFocus
                Timer.Interval = 50
            Case 16 To 17
                Num1.ForeColor = Num1.ForeColor - &H60606
                Num2.ForeColor = Num2.ForeColor - &H60606
                Midd.ForeColor = Midd.ForeColor - &H60606
                Ans.BackColor = Ans.BackColor - &H60606
            Case 18
                Invisiblize Num1
                Invisiblize Num2
                Invisiblize Midd
                Invisiblize Ans
                Timer.Interval = 100
                EndTms = EndTms + 1
            Case 19 To 24
                Me.BackColor = Me.BackColor + &H40404
                Hexa1.SetBackColor Me.BackColor
                Sleep Timer.Interval
            Case 25
                EndTms = EndTms + 1
                Me.BackColor = vbWhite
                Invisiblize Hexa1
                Hexa1.Peace
                With Midd
                    .Caption = "You Win"
                    .FontSize = 91
                    .Top = (Me.ScaleHeight - .Height) / 2
                    .Left = (Me.ScaleWidth - .Width) / 2
                    .BackColor = vbWhite
                    .ForeColor = vbWhite
                    .Visible = True
                End With
                Timer.Interval = 50
            Case 26 To 29
                Midd.ForeColor = Midd.ForeColor - &H30303
            Case 30
                Me.Caption = "You Have Become a Calculater. "
                Timer = False
        End Select
        EndTms = EndTms + Timer.Interval / 1000
    Else
        Time = Time - Timer.Interval / 1000
        If Time <= 0 Then
            NowScore = NowScore - Int(MaxTime / 3)
            If Score - NowScore > 20 Then
                NowScore = Score - 20
            End If
            If NowScore < 0 Then
                NowScore = 0
            End If
            Me.Caption = "目前难度系数：" & NowScore
            With Ans
                .Text = (Val(Num1) Mod Val(Num2)) & " "
                .BackColor = Me.BackColor
                .Alignment = 2
                .Locked = True
                .ForeColor = vbGreen
                .SelLength = 0
                .SelStart = Len(.Text)
            End With
            PadLet "累计失误次数：" & (Die + 1)
            DoEvents
            Sleep 1000
            DoEvents
            With Ans
                .Locked = False
                .Alignment = 0
                .BackColor = vbWhite
                .ForeColor = 0
            End With
            Die = Die + 1
            NXT
        End If
        Jdt.RenewProgress Time / MaxTime * 100
        If Music.Controls.currentPosition = 0 Then
            Silence = Silence + 1
            If Silence >= 5 Then
                Music.URL = App.Path & "\sound\m" & Int(Rnd * 3 + 1) & ".mp3"
                Music.Controls.play
            End If
        Else
            Silence = 0
        End If
    End If
End If
End Sub

Private Sub Si()
Timer = False
Hexa1.SetBackColor 0
Me.BackColor = 0
Me.Caption = MyName
Invisiblize Num1
Invisiblize Num2
Invisiblize Midd
Invisiblize Jdt
Invisiblize Ans
Invisiblize HS
Invisiblize STG
Restart.Visible = True
With Board
    .Caption = "分数：" & Score
    .FontSize = 65
    .Left = Midd.Left + (Midd.Width - .Width) / 2
End With
With Pad
    .Left = Midd.Left + (Midd.Width - .Width) / 2
    .Caption = "历史最高分：" & HighScore
    .FontSize = 38
    .Top = Board.Height + (Me.ScaleHeight - Board.Height - Pad.Height) * 0.6
    .ForeColor = vbWhite
    .BackColor = 0
End With
If Score > HighScore Then
    SaveSetting "iGlope", "MOD", "highscore", Score
    MsgBox "新纪录！", vbInformation + vbOKOnly
    Me.Caption = "新纪录：" & Score
End If
Music.Controls.stop
End Sub

Private Sub Invisiblize(Thing As Control)
Thing.Visible = False
End Sub

Private Sub Reboot()
Shell App.Path & "\" & App.EXEName, vbNormalFocus
Unload Me
End
End Sub

Sub WavPlay(strFileName As String)
If Fx Then
    If InStr(strFileName, " ") = 1 Then
        strFileName = Mid(strFileName, 2, 99)
    End If
    sndPlaySound App.Path & "\sound\" & strFileName & ".wav", 1
End If
End Sub

Private Sub ToggleFx_Click()
Fx = Not Fx
SaveSetting "iGlope", "MOD", "fx", Fx
If Not Fx Then
    sndPlaySoundStop 0, 0
End If
End Sub

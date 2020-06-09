VERSION 5.00
Begin VB.Form MainMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "我去冒险"
   ClientHeight    =   5916
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11484
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   22.2
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainMenu2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5916
   ScaleWidth      =   11484
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer LGM 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   7800
      Top             =   2760
   End
   Begin VB.Timer Main 
      Left            =   6960
      Top             =   3600
   End
   Begin VB.Line G 
      BorderColor     =   &H00000000&
      BorderWidth     =   16
      Index           =   0
      X1              =   -96
      X2              =   4000
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bt"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   25.8
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   666
      Index           =   0
      Left            =   4559
      TabIndex        =   0
      Top             =   2866
      Width           =   2366
   End
   Begin VB.Image Logo 
      Height          =   1368
      Index           =   3
      Left            =   7638
      Picture         =   "MainMenu2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1368
   End
   Begin VB.Image Logo 
      Height          =   1368
      Index           =   2
      Left            =   5918
      Picture         =   "MainMenu2.frx":330F
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1368
   End
   Begin VB.Image Logo 
      Height          =   1368
      Index           =   1
      Left            =   4198
      Picture         =   "MainMenu2.frx":8D2A
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1368
   End
   Begin VB.Image Logo 
      Height          =   1368
      Index           =   0
      Left            =   2478
      Picture         =   "MainMenu2.frx":ADEA
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1368
   End
   Begin VB.Line S 
      Index           =   0
      Visible         =   0   'False
      X1              =   4440
      X2              =   6000
      Y1              =   3120
      Y2              =   2520
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowCursor Lib "user32" (ByVal BShow As Long) As Long

Const LGfs As Double = 0.98
Const LGT As Long = 366
Dim LGx As Long, LGy As Long

Dim CFM As Boolean
Const CFMfs As Double = 0.98
Const CFMT As Long = 266
Dim CFMx As Long, CFMy As Long

Const Pi As Double = 3.14159265358979
Dim nM As Boolean
Dim k As Long, KK As Integer
Dim SX As Long, Sy As Long
Dim PX As Integer, PY As Integer
Const FxLays As Integer = 6
Dim FxNum(0 To 6) As Integer
Dim GB As Integer
Dim cy As Integer

Private Sub Bt_Click(Index As Integer)
Select Case Index
    Case 0  '开始
        LvChser.Show
        nM = True
        Unload Me
    Case 1  '神器
        SQ.Show
        nM = True
        Unload Me
    Case 2  '秘籍
        MiJi.Show
        nM = True
        Unload Me
    Case 3  '更新
        GX1.Show
        nM = True
        Unload Me
    Case 6
        Bt_Click 0
End Select
End Sub

Private Sub Bt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
G(0).BorderColor = vbBlue
G(1).BorderColor = vbBlue
End Sub

Private Sub Bt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
G(0).BorderColor = 0
G(1).BorderColor = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim pS(), Ps2(), PreGB As Integer
PreGB = GB
pS = Array(1, 2, 3, 0, 9, 9, 0)
Ps2 = Array(3, 0, 1, 2, 9, 9, 0)
Select Case KeyCode
    Case 38 'Up
        GB = Ps2(GB)
        cy = Bt(GB).Top + Bt(GB).Height / 2
        Sy = -166
    Case 40 'Down
        GB = pS(GB)
        cy = Bt(GB).Top + Bt(GB).Height / 2
        Sy = 166
    Case 13
        G(0).BorderColor = vbBlue
        G(1).BorderColor = vbBlue
        Exit Sub
    Case Else
        Exit Sub
End Select
If PreGB <> 6 Then
    Bt(PreGB).BackStyle = 0
    Bt(PreGB).ForeColor = 0
End If
Bt(GB).ForeColor = vbWhite
Bt(GB).BackStyle = 1
Bt(GB).BackColor = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 27
        Unload Me
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    G(0).BorderColor = 0
    G(1).BorderColor = 0
    Bt_Click GB
    Exit Sub
End If
End Sub

Private Sub Form_Load()
If GetSetting("iGlope", "woqu", "run", 0) = 0 Then
    MsgBox "请使用正常的方式启动游戏。", vbOKOnly + vbCritical, "警告"
    Unload Me
    End
End If
If Val(Command) >= 1 Then
    SQMsgBox.Show
    SQMsgBox.SQ_ID = Val(Command)
    nM = True
    Unload Me
    Exit Sub
End If
Randomize
Dim Ran
Ran = Rnd * 2 * Pi - Pi
SX = Sin(Ran) * 666
Sy = Cos(Ran) * 666
FxNum(0) = 1
Dim N As Integer
N = 36
For k = 1 To FxLays
    FxNum(k) = FxNum(k - 1) + N
    N = N * 0.26
Next k
Music.Visible = False
Music.ModeSwitch "menu"
Main.Interval = 2
Dim KK As Integer
For KK = 1 To FxLays
    For k = FxNum(KK - 1) To FxNum(KK) - 1
        Load S(k)
        With S(k)
            .BorderColor = RGB(0, 0, Int(Rnd() ^ 2 * 256))
            .BorderWidth = KK + 1
            .X1 = Rnd * Width
            .X2 = .X1
            .Y1 = Rnd * Height
            .Y2 = .Y1
            .Visible = True
        End With
    Next k
Next KK
Dim BtC
BtC = Array("开始游戏", "神　　器", "使用秘籍", "检查更新")
For k = 0 To 3
    With Bt(k)
    If k <> 0 Then
        Load Bt(k)
        .Top = Bt(k - 1).Top + .Height
        .Visible = True
    End If
    .Caption = BtC(k)
    End With
Next k
Load G(1)
With G(1)
    .X2 = Me.Width + 100
    .X1 = Me.Width - G(0).X2
    .Visible = True
    .ZOrder 0
End With
GB = 0
cy = Bt(0).Top + Bt(0).Height / 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (PX - X) ^ 2 + (PY - Y) ^ 2 <= 666666 Then
    SX = SX + (X - PX) / 16
    Sy = Sy + (Y - PY) / 16
End If
PX = X
PY = Y
If GB <> 7 Then
    If Y >= Bt(0).Top And Y <= Bt(3).Top + Bt(3).Height Then
        Dim Ind As Integer
        Ind = (Y - Bt(0).Top) \ Bt(0).Height
        If GB <> Ind Then
            If GB <> 6 Then
                Bt(GB).BackStyle = 0
                Bt(GB).ForeColor = 0
            End If
            GB = Ind
            Bt(GB).ForeColor = vbWhite
            Bt(GB).BackStyle = 1
            Bt(GB).BackColor = 0
            cy = Bt(GB).Top + Bt(GB).Height / 2
        End If
    Else
        If GB <> 6 Then
            Bt(GB).BackStyle = 0
            Bt(GB).ForeColor = 0
            GB = 6
        End If
        cy = Y
    End If
End If
End Sub

Private Sub LGM_Timer()
With Logo(2)
    .Left = .Left + LGx
    .Top = .Top + LGy
    If .Left <= 0 Or .Left + .Width >= Me.Width Then
        LGx = -LGx
        .Left = .Left + LGx
    End If
    If .Top <= 0 Or .Top + .Height >= Me.Height Then
        LGy = -LGy
        .Top = .Top + LGy
    End If
    LGx = LGx * LGfs
    LGy = LGy * LGfs
End With
If CFM Then
    With Me
        .Left = .Left + CFMx
        .Top = .Top + CFMy
        If .Left <= 0 Or .Left + .Width >= Screen.Width Then
            CFMx = -CFMx
            .Left = .Left + CFMx
        End If
        If .Top <= 0 Or .Top + .Height >= Screen.Height Then
            CFMy = -CFMy
            .Top = .Top + CFMy
        End If
        CFMx = CFMx * CFMfs
        CFMy = CFMy * CFMfs
        If Abs(CFMx) <= 40 Then
            CFMx = 0
        End If
        If Abs(CFMy) <= 40 Then
            CFMy = 0
        End If
    End With
End If
End Sub

Private Sub Logo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X + Logo(Index).Left, Y + Logo(Index).Top
End Sub

Private Sub Bt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X + Bt(Index).Left, Y + Bt(Index).Top
End Sub

Private Sub Logo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 2 Then
    If Button = 2 Then
        With Logo(2)
            '"6次"成就！
            .Tag = Val(.Tag) + 1
            Select Case Val(.Tag)
                Case 6
                    G(0).Visible = False
                    G(1).Visible = False
                    GB = 7
                    MsgBox "才单击六次就给你一个成就？这显然不合适。", vbExclamation + vbOKOnly, "我去冒险温馨地提示您："
                    MsgBox "成就应该是付出努力争取得来的。你至少要右键单击二十次，我才会把成就给你。", vbOKOnly + vbInformation, "我去冒险对你的深刻教育"
                Case 26
                    MsgBox "不，这样并不对。"
                    MsgBox "你刚才的二十次单击是如此的轻松！"
                    MsgBox "这个成就不能就这样白白送走！", , "没错，绝对不能！"
                    MsgBox "我觉得，你应该再用右键点它50下。", vbInformation, "只有这样才能体现出成就的价值！"
                Case 76
                    MsgBox "不——我依然没有感到满足！", vbCritical, "我去冒险渴望更多"
                    MsgBox "我一直认为，一个好的游戏作品，应该要去体现人类对梦想的追逐！", , "我一直以来的愿望。"
                    MsgBox "只有引导玩家去追逐他的梦想，一个游戏才能从一个游戏升华为一件艺术品！", , "你恍然大悟。"
                    .Top = Me.Height * 0.1
                    .Left = Me.Width * 0.1
                    Music.ModeSwitch "boss"
                Case 77
                    .Top = Me.Height * 0.5
                    .Left = Me.Width * 0.8
                Case 78
                    .Top = Me.Height * 0.8
                    .Left = Me.Width * 0.3
                Case 79
                    .Top = Me.Height * 0.3
                    .Left = Me.Width * 0.6
                Case 80
                    .Top = Me.Height * 0.8
                    .Left = Me.Width * 0.7
                Case 81
                    .Top = Me.Height * 0.1
                    .Left = Me.Width * 0.1
                Case 82
                    MsgBox "啊哈！你发现没有，这样一来游戏有趣多了！", , ""
                    MsgBox "让我们趁热打铁！现在，我会把你的鼠标隐藏，看看你能不能继续完成挑战！", , "一不做二不休:-)"
                    ShowCursor False
                    .Top = Me.Height * 0.8
                    .Left = Me.Width * 0.5
                    Bt(0).Visible = False
                    Bt(1).Visible = False
                    Bt(2).Visible = False
                    Bt(3).Visible = False
                Case 83
                    .Top = Me.Height * 0.1
                    .Left = Me.Width * 0.3
                Case 84
                    .Top = Me.Height * 0.3
                    .Left = Me.Width * 0.5
                Case 85
                    .Top = Me.Height * 0.66
                    .Left = Me.Width * 0.66
                Case 86
                    .Top = Me.Height * 0.1
                    .Left = Me.Width * 0.8
                Case 87
                    ShowCursor True
                    MsgBox "这真是太棒了！", , "我去冒险很开心"
                    MsgBox "我猜，你正深深地为这个游戏精妙的构思所折服！"
                    MsgBox "让我们再加把劲！"
                    LGM = True
                    .Top = Me.Height / 2
                    .Left = Me.Height / 2
                Case 88 To 96
                    If Rnd() <= 0.5 Then
                        LGx = LGT
                    Else
                        LGx = -LGT
                    End If
                    If Rnd() <= 0.5 Then
                        LGy = LGT
                    Else
                        LGy = -LGT
                    End If
                Case 97
                    MsgBox "再来点动作更大的吧！"
                    LGM = True
                    CFM = True
                    If Rnd() <= 0.5 Then
                        CFMx = CFMT
                    Else
                        CFMx = -CFMT
                    End If
                    If Rnd() <= 0.5 Then
                        CFMy = CFMT
                    Else
                        CFMy = -CFMT
                    End If
                Case 98 To 106
                    If Rnd() <= 0.5 Then
                        CFMx = CFMT
                    Else
                        CFMx = -CFMT
                    End If
                    If Rnd() <= 0.5 Then
                        CFMy = CFMT
                    Else
                        CFMy = -CFMT
                    End If
                    If Rnd() <= 0.5 Then
                        LGx = LGT
                    Else
                        LGx = -LGT
                    End If
                    If Rnd() <= 0.5 Then
                        LGy = LGT
                    Else
                        LGy = -LGT
                    End If
                Case 107
                    MsgBox "这真是太棒了！", vbInformation, "玩疯了！"
                    MsgBox "拿去吧。它属于你。", vbExclamation, "你值得拥有"
                    Shell App.Path & "\" & App.EXEName & ".exe 2", vbNormalFocus
                    Unload Me
                    End
            End Select
        End With
    End If
End If
End Sub

Private Sub Main_Timer()
Dim TarP
If KK Mod 3 = 0 Then
    For k = 0 To 1
        TarP = cy * 0.4 + G(k).Y1 * 0.6
        If Abs(G(k).Y1 - TarP) > 6 Then
            G(k).Y1 = TarP
        Else
            G(k).Y1 = cy
        End If
        G(k).Y2 = G(k).Y1
    Next k
End If
SX = SX * 0.99
Sy = Sy * 0.99
KK = KK Mod FxLays
KK = KK + 1
For k = FxNum(KK - 1) To FxNum(KK) - 1
    With S(k)
        .X1 = .X1 + SX * KK
        If .X1 > Width + 100 Then .X1 = -100: .Y1 = Rnd * Height
        If .X1 < -100 Then .X1 = Width + 100: .Y1 = Rnd * Height
        .Y1 = .Y1 + Sy * KK
        If .Y1 > Height + 100 Then .Y1 = -100: .X1 = Rnd * Width
        If .Y1 < -100 Then .Y1 = Height + 100: .X1 = Rnd * Width
        .X2 = .X1 - SX * KK
        .Y2 = .Y1 - Sy * KK
    End With
Next k
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not nM Then
    '游戏退出
    SaveSetting "iGlope", "woqu", "run", "0"
    If SQMsgBox.SQ_ID = 0 Then
        End
    End If
End If
End Sub


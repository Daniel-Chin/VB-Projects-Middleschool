VERSION 5.00
Begin VB.Form LvBoss 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "S.S.A.T."
   ClientHeight    =   6576
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8664
   Icon            =   "LvBoss.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6576
   ScaleWidth      =   8664
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox WTG 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   840
      ScaleHeight     =   492
      ScaleWidth      =   1212
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Stanley 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1920
      ScaleHeight     =   492
      ScaleWidth      =   1212
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Timer Plot 
      Index           =   1
      Left            =   4320
      Top             =   3840
   End
   Begin VB.Timer LXer 
      Interval        =   18
      Left            =   5520
      Top             =   3600
   End
   Begin 我去冒险.LinXin Fuck 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   855
      _ExtentX        =   1503
      _ExtentY        =   1926
   End
   Begin VB.TextBox Focuser 
      Height          =   370
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   -1006
      Width           =   850
   End
   Begin VB.Timer Killer 
      Left            =   6480
      Top             =   3240
   End
   Begin VB.Timer Drawer 
      Left            =   5880
      Top             =   4440
   End
   Begin VB.Timer Main 
      Left            =   5040
      Top             =   4440
   End
   Begin VB.Label Dont 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "别点"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   768
      Left            =   6636
      MouseIcon       =   "LvBoss.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   555
      Left            =   3795
      MouseIcon       =   "LvBoss.frx":1114
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Shape P 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      FillColor       =   &H000000FF&
      Height          =   370
      Left            =   6480
      Top             =   2160
      Width           =   970
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   6  'Cross
      Height          =   585
      Index           =   0
      Left            =   3188
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label PAD 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "LEVEL"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3233
      TabIndex        =   2
      Top             =   600
      Width           =   2205
   End
   Begin VB.Label Ad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "警告"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1290
      TabIndex        =   4
      Top             =   2640
      Width           =   825
   End
   Begin VB.Menu TopLeft 
      Caption         =   "菜单"
      Begin VB.Menu Quit 
         Caption         =   "返回主菜单"
      End
      Begin VB.Menu Suicide 
         Caption         =   "自杀"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "LvBoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'单字母变量占用表：g,f,w。k是循环变量，B是草稿
Option Explicit

Dim NearWin As Boolean
Dim ZuoEd As Boolean
Dim CuX As Long, CuY As Long
Dim Tms As Integer
Dim DiYiCi As Boolean
Dim LXV As Long
Dim Lv As Byte
Const YuanDian As Long = 7006, Wow As Long = 6
Const G As Long = 1 'd
Const F As Long = 6 'd
Dim JumpF As Long  'd

Private Type Walls
    Exist As Boolean
    XX As Long
    YY As Long
    Gao As Long
    Pang As Long
End Type
Private Type XiaoRer '小人儿
    XX As Long
    YY As Long
    Gao As Long
    Pang As Long
    Speed_Ver As Long '垂直速度
End Type
Private Type Keys
    Zuo As Boolean
    You As Boolean
    Jump As Boolean
End Type

Dim Wall(1 To 255) As Walls, WASD As Keys, Player As XiaoRer

Dim kTms As Byte

Private Sub Ad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X + Ad.Left, Y + Ad.Top
End Sub

Private Sub Bt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lv = 7 And Plot(7) Then
    LX
    Bt.Visible = False
End If
Bt.BorderStyle = 1
If Lv <= 5 Then
    Plot(Lv) = True
End If
Tms = 0
Select Case Lv
    Case 3
        Stanley.Visible = True
        Ad.Visible = False
End Select
End Sub

Private Sub Bt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ZuoEd Then
    LX
    Bt.Visible = False
End If
Bt.BackColor = RGB(166, 166, 166)
Bt.ForeColor = 0
End Sub

Private Sub Bt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Plot(Lv) And Not Plot(6) Then
    With Bt
        .BorderStyle = 0
        .Visible = False
    End With
Else
    If Lv = 7 Or 6 Then
        Plot(Lv) = True
    End If
End If
End Sub

Private Sub Dont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LX
End Sub

Private Sub Drawer_Timer()
With P
    If P.Left <> Player.XX * Wow Or P.Top <> YuanDian - (Player.YY + Player.Gao) * Wow Then
        .Left = Player.XX * Wow
        .Width = Player.Pang * Wow
        .Top = YuanDian - (Player.YY + Player.Gao) * Wow
        .Height = Player.Gao * Wow
    End If
    If NearWin Then
        Dim bX As Long, bY As Long
        bX = .Left + .Width / 2
        bY = .Top + .Height / 2
        With Fuck
            If bX >= .Left And bX <= .Left + .Width And bY >= .Top And bY <= .Top + .Height Then
                Win
            End If
        End With
    End If
End With
Drawer.Interval = Main.Interval
End Sub

Private Sub Focuser_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub Focuser_KeyUp(KeyCode As Integer, Shift As Integer)
Form_KeyUp KeyCode, Shift
End Sub

Private Sub Focuser_LostFocus()
LX
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 38
        WASD.Jump = True
        If Lv = 7 And JumpF = 26 Then
            NearWin = True
        End If
    Case 37
        WASD.Zuo = True
    Case 39
        WASD.You = True
End Select
If Killer.Interval = 0 Then Main.Interval = 18
LX
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LX
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bt.BackColor = 0
Bt.ForeColor = vbWhite
CuX = X
CuY = Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell App.Path & "\" & App.EXEName, vbNormalFocus
End
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 38
        WASD.Jump = False
    Case 37
        WASD.Zuo = False
    Case 39
        WASD.You = False
End Select
If DiYiCi Then
    DiYiCi = False
Else
    LX
End If
End Sub

Private Sub Form_Load()
Music.ModeSwitch "boss"
Dim k As Byte
Height = YuanDian
With Wall(1)
    .Exist = True
    .XX = W(0).Left / Wow
    .YY = (YuanDian - W(0).Top - W(0).Height) / Wow
    .Gao = W(0).Height / Wow
    .Pang = W(0).Width / Wow
End With
For k = 2 To 7
    Load Plot(k)
Next k
Reset
Main.Interval = 18
Drawer.Interval = 18
Drawer_Timer
Me.Caption = "我去冒险 - " & Me.Name
End Sub

Private Function Qiang(Optional XX As Long = -666, Optional YY As Long = -666) As Byte
If XX = -666 Then XX = Player.XX
If YY = -666 Then YY = Player.YY
Dim A(1 To 4) As Boolean
Dim k As Byte
k = 255
Do Until k = 0
    With Wall(k)
        If .Exist Then
            A(1) = XX + Player.Pang >= .XX
            A(2) = XX <= .XX + .Pang
            A(3) = YY + Player.Gao >= .YY
            A(4) = YY <= .YY + .Gao
            If A(1) And A(2) And A(3) And A(4) Then
                Exit Do
            End If
        End If
    End With
    k = k - 1
Loop
Qiang = k
'限制边框
If XX <= 0 Or XX + Player.Pang >= Width / Wow Then Qiang = 255
End Function

Private Sub Killer_Timer()
If kTms = 10 Then
    Killer.Interval = 0
    Si True
Else
    kTms = kTms + 1
    P.FillStyle = 1 - P.FillStyle
End If
End Sub

Private Sub LinXin_GotFocus()
Focuser.SetFocus
End Sub

Private Sub LXer_Timer()
If LXV = Wow * 2 Then Wall(1).Exist = False
LXV = LXV + Wow
W(0).Top = W(0).Top + LXV
Main.Interval = 18
End Sub

Private Sub Main_Timer()
Dim B As Long
If Qiang() Then
    Si
Else
    With Player
        Dim DongLe As Boolean
        B = .XX + (WASD.Zuo - WASD.You) * F
        Do While Qiang(B)
            B = B + WASD.You - WASD.Zuo
        Loop
        If .XX <> B Then DongLe = True
        .XX = B
        '横如上
        '纵如下
        Dim Grnd As Boolean
        Grnd = isGrnd()
        If Grnd Then
            If WASD.Jump Then
                .Speed_Ver = JumpF
            Else
                .Speed_Ver = 0
            End If
        Else
            .Speed_Ver = .Speed_Ver - G
        End If
        B = .YY + .Speed_Ver
        If .Speed_Ver > 0 Then
            '顶踩事件！！！
            If Qiang(, B) Then
                .Speed_Ver = 0
                Do While Qiang(, B)
                    B = B - 1
                Loop
            End If
        Else
            If Qiang(, B) Or Grnd Then
                .Speed_Ver = 0
                Do While Qiang(, B)
                    B = B + 1
                Loop
            End If
        End If
        If .YY <> B Then DongLe = True
        .YY = B
        If .YY <= -166 Then Si True
    End With
End If
If Not DongLe And Grnd Then
    Main.Interval = 0
End If
Drawer.Interval = Main.Interval
End Sub

Private Sub Si(Optional TouTouDe As Boolean = False)
If Not TouTouDe Then
    Main.Interval = 0
    Drawer.Interval = 0
    TouTouDe = True
    Killer.Interval = 106
    kTms = 0
    P.FillColor = &HFF&
    Drawer_Timer
Else
    MsgBox "你死了。", vbCritical + vbOKOnly, Caption
    Reset
End If
End Sub

Private Function isGrnd() As Boolean
Debug.Assert Qiang() = 0
Dim A(1 To 3) As Boolean
Dim k As Byte
k = 255
Do Until k = 0
    With Wall(k)
        If .Exist Then
            A(1) = Player.XX + Player.Pang >= .XX
            A(2) = Player.XX <= .XX + .Pang
            A(3) = Player.YY = .YY + .Gao + 1
            If A(1) And A(2) And A(3) Then
                isGrnd = True
                Exit Do
            End If
        End If
    End With
    k = k - 1
Loop
End Function

Private Sub Reset()
With Player
    .Speed_Ver = 0
    .XX = 666 'd
    .YY = 366 'd
    .Gao = 80 'd
    .Pang = 60 'd
End With
With WASD
    .Jump = False
    .You = False
    .Zuo = False
End With
Main.Interval = 18
Drawer.Interval = 18
With P
    .FillStyle = 0
    .FillColor = &HFF0000
End With
JumpF = 26
Bt.BorderStyle = 0
Lv = 1
SetLv
LXer = False
W(0).Top = 5520
Wall(1).Exist = True
DiYiCi = True
Dim k As Byte
For k = 1 To 7
    Plot(k).Interval = 18
    Plot(k) = False
Next k
Fuck.Top = 0
Fuck.Left = 0
Bt.Left = 3795
Bt.Visible = True
With Stanley
    .Visible = False
    .Width = 6666
    .Height = 4666
    .Top = (Height - .Height) / 2 - 666
    .Left = (Width - .Width) / 2
    .PaintPicture LoadPicture(App.Path & "\ad.jpg"), 0, 0, .Width, .Height
End With
Ad.Visible = True
Dont.Visible = False
WTG.Visible = False
ZuoEd = False
NearWin = False
Music.ModeSwitch "boss"
End Sub

Private Sub PAD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X + Pad.Left, Y + Pad.Top
End Sub


Private Sub Plot_Timer(Index As Integer)
Tms = Tms + 1
With Fuck
    Select Case Index
        Case 1
            If Tms >= 424 Then
                Fuck.Top = Fuck.Top + Wow
                If Tms >= 500 Then
                    Lv = 2
                    Plot(1) = False
                    SetLv
                End If
            Else
                Fuck.Left = Fuck.Left + 3 * Wow
            End If
        Case 2
            Select Case Tms
                Case 1 To 212
                    .Left = .Left - 3 * Wow
                Case 250 To 499
                    .Top = .Top + 2 * Wow
                Case 736 To 900
                    .Top = .Top - 3 * Wow
                Case 901 To 1090
                    .Left = .Left - 3 * Wow
                Case 1091 To 1170
                    .Top = .Top + Wow
                Case 1191
                    Lv = 3
                    Plot(2) = False
                    SetLv
            End Select
        Case 3
            Select Case Tms
                Case 199
                    Stanley.Visible = False
                    Ad.Visible = True
                Case 200 To 600
                    .Left = .Left + 3 * Wow
                Case 601 To 666
                    .Top = .Top + Wow
                Case 667
                    Lv = 4
                    Plot(3) = False
                    SetLv
            End Select
        Case 4
            Select Case Tms
                Case 1 To 186
                    .Left = .Left - 3 * Wow
                    If Tms <= 56 Then
                        Player.XX = Player.XX + 2
                    Else
                        Player.XX = Player.XX + 1
                    End If
                    Drawer_Timer
                Case 187 To 424
                    .Left = .Left - 3 * Wow
                Case 425 To 500
                    .Top = .Top + Wow
                Case 501
                    Lv = 5
                    Plot(4) = False
                    SetLv
            End Select
        Case 5
            Select Case Tms
                Case 1
                    Dont.Visible = True
                Case 2 To 424
                    .Left = .Left + 3 * Wow
                    Dim B As Integer
                    B = Sqr((CuX - Dont.Left - Dont.Width / 2) ^ 2 + (CuY - Dont.Top - Dont.Height / 2) ^ 2) / 40
                    Dont.ForeColor = RGB(B, Abs(B - 128), B)
                    Dont.BackColor = RGB(255 - B, 255 - B, 255 - B)
                Case 425 To 500
                    .Top = .Top + Wow
                Case 501
                    Dont.Visible = False
                    Lv = 6
                    Plot(5) = False
                    SetLv
            End Select
        Case 6
            Select Case Tms
                Case 1
                    Bt.Left = Bt.Left - 1666
                    Bt.BackColor = 0
                    Bt.ForeColor = vbWhite
                    Bt.Visible = True
                    ZuoEd = True
                Case 166
                    Bt.Visible = False
                    Bt.Left = Bt.Left + 1666
                    ZuoEd = False
                Case 167 To 591
                    .Left = .Left - 3 * Wow
                Case 592
                    Lv = 7
                    Plot(6) = False
                    SetLv
            End Select
        Case 7
            Select Case Tms
                Case 106
                    With Bt
                        .BorderStyle = 0
                        .Visible = False
                    End With
                Case 107 To 531
                    .Left = .Left + 3 * Wow
                Case 532 To 608
                    .Top = .Top + Wow * 2
                Case 609
                    With WTG
                        .AutoRedraw = True
                        .Cls
                        .Height = 4666
                        .Width = 6666
                        .Left = (Me.Width - .Width) / 2
                        .Top = (Me.Height - .Height) / 2
                        .PaintPicture LoadPicture(App.Path & "\wtg.jpg"), 0, 0, .Width, .Height
                        .Visible = True
                    End With
                Case 866
                    WTG.Visible = False
                    JumpF = 26
                Case 867 To 1079
                    .Left = .Left - Wow * 2
                Case 1166 To 1265
                    .Top = .Top - 6 * Wow
                Case 1366 To 1367
                    Tms = 1366
            End Select
    End Select
End With
End Sub

Private Sub Quit_Click()
Unload Me
End Sub

Private Sub Stanley_GotFocus()
Focuser.SetFocus
End Sub

Private Sub Stanley_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= 5000 And Y <= 444 Then
    LX
    Stanley.Visible = False
Else
    Shell "explorer http://www.stanley-house.com/index_cn.html", vbMaximizedFocus
End If
End Sub

Private Sub Suicide_Click()
Si
End Sub

Private Sub Win()
MsgBox "你终于打通了Boss关卡！", vbOKOnly + vbInformation, Me.Caption
Shell App.Path & "\" & App.EXEName & ".exe 1", vbNormalFocus
Music.ModeSwitch "menu"
Sleep 4000
Shell App.Path & "\zimu.exe", vbNormalFocus
MsgBox "你解锁了奖励关卡。快去主菜单看看吧！", vbInformation + vbOKOnly, "奇怪的小冒险"
Unload Me
End Sub

Private Sub LX() '坠落的空岛
If LXer = 0 Then
    LXer = True
    LXV = 0
End If
End Sub

Private Sub SetLv()
Pad = "STAGE " & Lv
If Lv = 7 Then Pad = "STAGE FINAL"
Bt.Visible = True
Select Case Lv
    Case 1
        Ad = "空岛不稳定" & String(2, Chr(10)) & "请勿大声喧哗"
        Ad.BackColor = vbWhite
    Case 5
        Ad = "空岛不稳定" & String(2, Chr(10)) & "请勿大声喧哗"
        Ad.BackColor = vbWhite
    Case 2
        JumpF = 7
    Case 4
        Ad = "天气预报" & String(2, Chr(10)) & "西风3.0级"
        Ad.BackColor = RGB(255, 66, 66)
End Select
End Sub

Private Sub WTG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
WTG.Visible = False
End Sub

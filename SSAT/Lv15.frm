VERSION 5.00
Begin VB.Form Lv15 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "S.S.A.T."
   ClientHeight    =   6576
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8664
   Icon            =   "Lv15.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6576
   ScaleWidth      =   8664
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame YJCD 
      BackColor       =   &H00E0E0E0&
      Caption         =   "右键菜单"
      Height          =   1690
      Left            =   6480
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   2050
      Begin VB.Label CD 
         Alignment       =   2  'Center
         Caption         =   "作者长的很帅！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1810
      End
      Begin VB.Label CD 
         Alignment       =   2  'Center
         Caption         =   "我去冒险"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1810
      End
      Begin VB.Label CD 
         Alignment       =   2  'Center
         Caption         =   "很坑爹吧？"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1810
      End
      Begin VB.Label CD 
         Alignment       =   2  'Center
         Caption         =   "加固地板"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1810
      End
   End
   Begin VB.TextBox Focuser 
      BackColor       =   &H00FFFFFF&
      Height          =   370
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   -1006
      Width           =   850
   End
   Begin 我去冒险.LinXin LinXin 
      Height          =   1210
      Left            =   7440
      TabIndex        =   1
      Top             =   3720
      Width           =   850
      _ExtentX        =   1503
      _ExtentY        =   2138
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
   Begin VB.Label Pad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "看看菜单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   3375
      TabIndex        =   8
      Top             =   1080
      Width           =   1930
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   6  'Cross
      Height          =   220
      Index           =   8
      Left            =   7440
      Top             =   4680
      Width           =   850
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "?"
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
      Height          =   490
      Left            =   2880
      TabIndex        =   0
      Top             =   2280
      Width           =   490
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      Height          =   460
      Index           =   7
      Left            =   2880
      Top             =   2280
      Width           =   490
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   6  'Cross
      Height          =   340
      Index           =   6
      Left            =   2040
      Top             =   4920
      Width           =   4330
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   6  'Cross
      Height          =   340
      Index           =   5
      Left            =   5760
      Top             =   3720
      Width           =   1210
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   6  'Cross
      Height          =   340
      Index           =   4
      Left            =   1440
      Top             =   3720
      Width           =   1210
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   6  'Cross
      Height          =   1660
      Index           =   3
      Left            =   6000
      Top             =   4080
      Width           =   730
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   6  'Cross
      Height          =   1660
      Index           =   2
      Left            =   1680
      Top             =   4080
      Width           =   730
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   6  'Cross
      Height          =   3220
      Index           =   1
      Left            =   6360
      Top             =   4920
      Width           =   2290
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
      Height          =   3220
      Index           =   0
      Left            =   -240
      Top             =   4920
      Width           =   2290
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
Attribute VB_Name = "Lv15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'单字母变量占用表：g,f,w。k是循环变量，B是草稿
Option Explicit

Dim nM As Boolean
Dim Tms As Byte, dTms As Byte
Const YuanDian As Long = 7006, Wow As Long = 6
Const G As Long = 1 'd
Const F As Long = 6 'd
Const JumpF As Long = 26 'd

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

Private Sub CD_Click(Index As Integer)
Select Case Index
    Case 0
        YJCD.Visible = False
        CD(0).Enabled = False
        Wall(7).Exist = True
    Case 3
        YJCD.Visible = False
        Tms = Tms + 1
        If Tms >= 6 Then
            MsgBox "这是我的qq：2208825156，有空联系我哦！", vbOKOnly + vbInformation, "男的就算了。"
            CD(3).Enabled = False
        Else
            MsgBox "谢谢，你也很帅！", vbOKOnly + vbInformation, "太机智了"
        End If
End Select
End Sub

Private Sub CD_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
CLB
If Index = 0 Or Index = 3 Then
    CD(Index).BackColor = &HFFFF00
End If
End Sub

Private Sub Drawer_Timer()
With P
    If P.Left <> Player.XX * Wow Or P.Top <> YuanDian - (Player.YY + Player.Gao) * Wow Then
        .Left = Player.XX * Wow
        .Width = Player.Pang * Wow
        .Top = YuanDian - (Player.YY + Player.Gao) * Wow
        .Height = Player.Gao * Wow
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 38
        WASD.Jump = True
    Case 37
        WASD.Zuo = True
    Case 39
        WASD.You = True
End Select
If Killer.Interval = 0 Then Main = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CLB
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    With YJCD
        .Top = Y
        .Left = X
        .Visible = Not .Visible
    End With
Else
    YJCD.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not nM Then
    '点红叉叉
    Shell App.Path & "\" & App.EXEName, vbNormalFocus
    End
Else
    Unload Lv14_5
End If
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
End Sub

Private Sub Form_Load()
Height = YuanDian
'加载墙壁
Dim k As Byte
For k = W.LBound To W.UBound
    With Wall(k + 1)
        .Exist = True
        .XX = W(k).Left / Wow
        .YY = (YuanDian - W(k).Top - W(k).Height) / Wow
        .Gao = W(k).Height / Wow
        .Pang = W(k).Width / Wow
    End With
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

Private Sub Label1_Click()
Form_MouseUp 1, 0, 0, 0
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CLB
End Sub

Private Sub LinXin_GotFocus()
Focuser.SetFocus
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
                '顶
                Select Case Qiang(, B + 1)
                    Case 8 '?
                        dTms = dTms + 1
                        If dTms <= 6 Then
                            Pad.Visible = True
                            Label1.Top = Label1.Top - 6 * Wow
                            Wall(8).YY = Wall(8).YY + 6
                            W(7).Top = Label1.Top
                        Else
                            Pad = "看看右键菜单"
                            If dTms = 255 Then
                                MsgBox "你太无聊了！", vbOKOnly + vbCritical
                                Wall(8).Exist = False
                                Label1.Visible = False
                                Wall(7).Exist = False
                            End If
                        End If
                End Select
            End If
        Else
            If Qiang(, B) Or Grnd Then
                .Speed_Ver = 0
                Do While Qiang(, B)
                    B = B + 1
                Loop
                '踩
                Select Case Qiang(, B - 1)
                    Case 9 'Win
                        If MsgBox("你死了。", vbOKCancel + vbCritical, "我去冒险") = vbOK Then
                            Si True, False
                        Else
                            Win
                        End If
                    Case 8
                        Si
                End Select
            End If
        End If
        If .YY <> B Then DongLe = True
        .YY = B
        If .YY <= -166 Then Si
    End With
End If
If Not DongLe And Grnd Then
    Main.Enabled = False
End If
Drawer.Interval = Main.Interval
End Sub

Private Sub Si(Optional TouTouDe As Boolean = False, Optional SP As Boolean = True)
If Not TouTouDe Then
    Main.Enabled = False
    Drawer.Enabled = False
    TouTouDe = True
    Killer.Interval = 106
    kTms = 0
    P.FillColor = &HFF&
    Drawer_Timer
Else
    If SP Then
        MsgBox "你死了。", vbCritical + vbOKOnly, Caption
    End If
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
    .XX = 56 'd
    .YY = 416 'd
    .Gao = 80 'd
    .Pang = 60 'd
End With
With WASD
    .Jump = False
    .You = False
    .Zuo = False
End With
Main = True
Drawer = True
With P
    .FillStyle = 0
    .FillColor = &HFF0000
End With
Wall(7).Exist = False
YJCD.Visible = False
CD(0).Enabled = True
Tms = 0
Pad.Visible = False
Pad = "看看主菜单"
dTms = 0
Wall(8).Exist = True
Label1.Visible = True
Label1.Top = 2280
W(7).Top = Label1.Top
Wall(8).YY = 711
End Sub

Private Sub Quit_Click()
MsgBox "你去查看了主菜单。没有发现任何可疑的东西。", vbInformation + vbOKOnly
End Sub

Private Sub Suicide_Click()
Si
End Sub

Private Sub Win()
If Val(GetSetting("iGlope", "woqu", "lv", 1)) < 19 Then
    SaveSetting "iGlope", "woqu", "lv", "19"
End If
MsgBox "恭喜你打通了" & Me.Name & "！按OK进入下一关。", vbOKOnly + vbInformation, Me.Caption
MsgBox "注意！下一关就是Boss关卡。" & Chr(13) & Chr(10) & "你准备好了吗？", vbDefaultButton3 + vbExclamation
LvBoss.Visible = True
nM = True
Unload Me
End Sub

Private Sub CLB()
CD(0).BackColor = &H8000000F
CD(3).BackColor = &H8000000F
End Sub

Private Sub YJCD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CLB
End Sub

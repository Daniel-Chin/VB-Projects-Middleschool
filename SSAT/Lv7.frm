VERSION 5.00
Begin VB.Form Lv7 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "S.S.A.T."
   ClientHeight    =   6552
   ClientLeft      =   96
   ClientTop       =   444
   ClientWidth     =   8688
   Icon            =   "Lv7.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6552
   ScaleWidth      =   8688
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Focuser 
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
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   6  'Cross
      Height          =   340
      Index           =   6
      Left            =   0
      Top             =   -1000
      Width           =   4330
   End
   Begin VB.Label YY 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "右"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   4605
      TabIndex        =   4
      Top             =   1320
      Width           =   850
   End
   Begin VB.Label ZZ 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "左"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   3225
      TabIndex        =   3
      Top             =   1320
      Width           =   850
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
   Begin VB.Label Pad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "现在鼠标也能控制人物了。"
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
      Left            =   1455
      TabIndex        =   0
      Top             =   720
      Width           =   5770
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      Height          =   460
      Index           =   7
      Left            =   5280
      Top             =   -800
      Width           =   490
   End
   Begin VB.Shape WW 
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
Attribute VB_Name = "Lv7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'单字母变量占用表：g,f,w。k是循环变量，B是草稿
Option Explicit

Dim nM As Boolean
Dim Draging As Boolean, DragX As Long, DragY As Long
Const YuanDian As Long = 7006, Wow As Long = 6
Const g As Long = 1 'd
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

Private Sub Drawer_Timer()
If Draging Then
    Main = False
    Drawer = False
End If
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim A(1 To 4) As Boolean
With P
    A(1) = X >= .Left
    A(2) = X <= .Left + .Width
    A(3) = Y >= .Top
    A(4) = Y <= .Top + .Height
End With
If A(1) And A(2) And A(3) And A(4) Then
    Draging = True
    DragX = X
    DragY = Y
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Draging Then
    P.Top = P.Top + Y - DragY
    P.Left = P.Left + X - DragX
    DragX = X
    DragY = Y
End If
Dim A(1 To 4) As Boolean
With LinXin
    A(1) = X >= .Left
    A(2) = X <= .Left + .Width
    A(3) = Y >= .Top
    A(4) = Y <= .Top + .Height
End With
If A(1) And A(2) And A(3) And A(4) Then Win
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not nM Then
    '点红叉叉
    Shell App.Path & "\" & App.EXEName, vbNormalFocus
    End
Else
    Unload Lv6
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

Private Sub LinXin_GotFocus()
Focuser.SetFocus
End Sub

Private Sub Main_Timer()
Dim B As Long
If Draging Then
    Main = False
    Drawer = False
Else
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
                .Speed_Ver = .Speed_Ver - g
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
                            '
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
                            Win
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
End If
End Sub

Private Sub Si(Optional TouTouDe As Boolean = False)
If Not TouTouDe Then
    Main.Enabled = False
    Drawer.Enabled = False
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
Draging = False
ZZ_MouseUp 1, 0, 0, 0
YY_MouseUp 1, 0, 0, 0
End Sub

Private Sub PAD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X + Pad.Left, Y + Pad.Top

End Sub

Private Sub Quit_Click()
Unload Me
End Sub

Private Sub Suicide_Click()
Si
End Sub

Private Sub Win()
If Val(GetSetting("iGlope", "woqu", "lv", 1)) < 9 Then
    SaveSetting "iGlope", "woqu", "lv", "9"
End If
MsgBox "恭喜你打通了" & Me.Name & "！按OK进入下一关。", vbOKOnly + vbInformation, Me.Caption
Lv8.Visible = True
nM = True
Unload Me
End Sub

Private Sub ZZ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If YY.BackColor = &HFFC0C0 Then
    ZZ.BackColor = ZZ.BackColor - RGB(100, 100, 20)
    With WASD
        .Jump = True
        .Zuo = True
    End With
    If Killer.Interval = 0 Then Main = True
End If
End Sub

Private Sub ZZ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X + ZZ.Left, Y + ZZ.Top
End Sub

Private Sub YY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X + YY.Left, Y + YY.Top
End Sub

Private Sub ZZ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ZZ.BackColor = &HFFC0C0 - RGB(100, 100, 20) Then
    ZZ.BackColor = ZZ.BackColor + RGB(100, 100, 20)
    With WASD
        .Jump = False
        .Zuo = False
    End With
End If
End Sub

Private Sub YY_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If YY.BackColor = &HFFC0C0 Then
    YY.BackColor = YY.BackColor - RGB(100, 100, 20)
    With WASD
        .Jump = True
        .You = True
    End With
    If Killer.Interval = 0 Then Main = True
End If
End Sub

Private Sub YY_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If YY.BackColor = &HFFC0C0 - RGB(100, 100, 20) Then
    YY.BackColor = YY.BackColor + RGB(100, 100, 20)
    With WASD
        .Jump = False
        .You = False
    End With
End If
End Sub


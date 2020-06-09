VERSION 5.00
Begin VB.Form LvBi 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "S.S.A.T."
   ClientHeight    =   6552
   ClientLeft      =   96
   ClientTop       =   144
   ClientWidth     =   8688
   Icon            =   "LvBi.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6552
   ScaleWidth      =   8688
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer FLer 
      Left            =   4080
      Top             =   3600
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
   Begin VB.Label FLSH 
      Height          =   372
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Cloud 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cloud"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   66
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1812
      Left            =   2538
      TabIndex        =   1
      Top             =   480
      Width           =   3612
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   0  'Solid
      Height          =   456
      Index           =   12
      Left            =   6480
      Top             =   2640
      Width           =   492
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   0  'Solid
      Height          =   456
      Index           =   11
      Left            =   5524
      Top             =   2640
      Width           =   492
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   0  'Solid
      Height          =   456
      Index           =   10
      Left            =   4572
      Top             =   2640
      Width           =   492
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   0  'Solid
      Height          =   456
      Index           =   8
      Left            =   2668
      Top             =   2640
      Width           =   492
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   0  'Solid
      Height          =   456
      Index           =   9
      Left            =   3620
      Top             =   2640
      Width           =   492
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   0  'Solid
      Height          =   456
      Index           =   7
      Left            =   1716
      Top             =   2640
      Width           =   492
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
      Height          =   372
      Left            =   7440
      Top             =   3240
      Width           =   972
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
End
Attribute VB_Name = "LvBi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'单字母变量占用表：g,f,w。k是循环变量，B是草稿
Option Explicit

Dim Bi(0 To 63) As Integer
Dim GB As Integer
Const GBTop As Integer = 63

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
If P.Left <> Player.XX * Wow Or P.Top <> YuanDian - (Player.YY + Player.Gao) * Wow Then
    With P
        .Left = Player.XX * Wow
        .Width = Player.Pang * Wow
        .Top = YuanDian - (Player.YY + Player.Gao) * Wow
        .Height = Player.Gao * Wow
    End With
End If
Drawer.Interval = Main.Interval
End Sub

Private Sub FLer_Timer()
FLSH.Visible = False
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
Me.Caption = "异"
FLSH.Move 0, 0, Me.Width, Me.Height
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

Private Sub Form_Unload(Cancel As Integer)
Shell App.Path & "\" & App.EXEName, vbNormalFocus
End
End Sub

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
                Debug.Assert Qiang(, B + 1) >= 8
                '8-13
                W(Qiang(, B + 1) - 1).FillStyle = 1 - W(Qiang(, B + 1) - 1).FillStyle
                Dim TempBi As Integer
                Dim EtT As Integer
                Dim tLG As Integer
                tLG = 1
                For EtT = 13 To 8 Step -1
                    TempBi = TempBi + tLG * W(EtT - 1).FillStyle
                    tLG = tLG * 2
                Next EtT
                Bi(GB) = TempBi
                For EtT = 0 To GB - 1
                    If Bi(EtT) = TempBi Then
                        FLer.Interval = 366
                        FLSH.BackColor = vbRed
                        FLSH.Visible = True
                        Si
                        Exit Sub
                    Else
                        FLer.Interval = 136
                        FLSH.BackColor = vbGreen
                        FLSH.Visible = True
                    End If
                Next EtT
                GB = GB + 1
                Cls
                Cloud = GBTop - GB
                If GB = GBTop Then
                    Win
                End If
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
        If .YY <= -166 Then Si
    End With
End If
If Not DongLe And Grnd Then
    Main.Enabled = False
End If
Drawer.Interval = Main.Interval
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
    MsgBox "第二次出现了相同的图案！", vbCritical + vbOKOnly, Caption
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
Wall(5).Exist = False
Wall(6).Exist = False
Wall(3).Exist = False
Wall(4).Exist = False
For GB = 0 To GBTop
    Bi(GB) = 0
Next GB
For GB = 7 To 12
    W(GB).FillStyle = 0
Next GB
GB = 1
Cloud = GBTop - 1
Music.ModeSwitch "boss"
End Sub

Private Sub Win()
SaveSetting "iGlope", "woqu", "j6", 1
MsgBox "你居然打通了奖励关卡Binary！不可思议啊！没看攻略吧？", vbOKOnly + vbInformation, "Unbelievable"
Shell App.Path & "\" & App.EXEName, vbNormalFocus
End
End Sub


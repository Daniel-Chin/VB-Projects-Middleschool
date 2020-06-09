VERSION 5.00
Begin VB.Form Lv14_5 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "S.S.A.T."
   ClientHeight    =   6552
   ClientLeft      =   96
   ClientTop       =   444
   ClientWidth     =   8688
   Icon            =   "Lv14.5.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6552
   ScaleWidth      =   8688
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox QRDo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   1086
      Left            =   4800
      ScaleHeight     =   1092
      ScaleWidth      =   2172
      TabIndex        =   7
      Top             =   600
      Width           =   2172
   End
   Begin VB.PictureBox QRUp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   1086
      Left            =   1440
      ScaleHeight     =   1092
      ScaleWidth      =   2172
      TabIndex        =   6
      Top             =   600
      Width           =   2172
   End
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
   Begin VB.Label WH 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
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
      Height          =   492
      Index           =   3
      Left            =   4800
      TabIndex        =   5
      Top             =   2760
      Width           =   492
   End
   Begin VB.Label WH 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
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
      Height          =   492
      Index           =   2
      Left            =   3840
      TabIndex        =   4
      Top             =   2760
      Width           =   492
   End
   Begin VB.Label WH 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
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
      Height          =   492
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      Height          =   456
      Index           =   11
      Left            =   4800
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      Height          =   456
      Index           =   10
      Left            =   3840
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      Height          =   456
      Index           =   9
      Left            =   2880
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      FillStyle       =   6  'Cross
      Height          =   220
      Index           =   7
      Left            =   7440
      Top             =   4680
      Width           =   850
   End
   Begin VB.Label WH 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
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
      Height          =   492
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape W 
      BackColor       =   &H00000000&
      BorderWidth     =   5
      Height          =   456
      Index           =   8
      Left            =   1920
      Top             =   2760
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
      Height          =   336
      Index           =   5
      Left            =   5760
      Top             =   2040
      Width           =   1212
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
      Height          =   3456
      Index           =   3
      Left            =   6000
      Top             =   2280
      Width           =   732
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
Attribute VB_Name = "Lv14_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'单字母变量占用表：g,f,w。k是循环变量，B是草稿
Option Explicit

Dim PassWd(0 To 3) As Integer
Dim DW As Long

Dim DragX As Long
Dim DragY As Long

Dim nM As Boolean
Dim FTms As Long
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
Main.Interval = 18
Drawer.Interval = 18
Drawer_Timer
Me.Caption = "我去冒险 - " & Me.Name
Dim QR(0 To 41)
QR(0) = Array(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 1, 0, 0, 0, 1, 1, 0, 1, 0, 1, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0)
QR(1) = Array(0, 1, 1, 1, 1, 1, 0, 1, 1, 0, 0, 1, 0, 1, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, 0, 1, 1, 0, 1, 1, 1, 1, 0, 1, 1, 1, 1, 1, 0)
QR(2) = Array(0, 1, 0, 0, 0, 1, 0, 1, 1, 1, 0, 0, 0, 0, 0, 1, 1, 0, 1, 1, 0, 1, 0, 1, 0, 1, 1, 1, 1, 0, 0, 0, 1, 1, 0, 1, 0, 0, 0, 1, 0)
QR(3) = Array(0, 1, 0, 0, 0, 1, 0, 1, 0, 0, 1, 0, 1, 1, 1, 0, 0, 1, 1, 0, 1, 0, 1, 1, 1, 1, 0, 0, 1, 1, 0, 1, 0, 1, 0, 1, 0, 0, 0, 1, 0)
QR(4) = Array(0, 1, 0, 0, 0, 1, 0, 1, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 1, 0, 0, 1, 1, 0, 1, 0, 0, 0, 0, 1, 0, 1, 0, 0, 0, 1, 0)
QR(5) = Array(0, 1, 1, 1, 1, 1, 0, 1, 1, 0, 1, 1, 0, 1, 1, 1, 0, 0, 1, 0, 1, 1, 0, 1, 0, 0, 1, 0, 0, 0, 1, 1, 1, 1, 0, 1, 1, 1, 1, 1, 0)
QR(6) = Array(0, 0, 0, 0, 0, 0, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0)
QR(7) = Array(1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 1, 0, 0, 1, 1, 0, 1, 0, 1, 0, 1, 1, 1, 0, 1, 1, 1, 1, 1, 1, 1, 1)
QR(8) = Array(1, 1, 0, 0, 1, 1, 0, 1, 1, 0, 1, 1, 1, 1, 1, 1, 0, 0, 0, 1, 0, 1, 1, 0, 1, 0, 1, 1, 1, 0, 0, 0, 1, 0, 0, 1, 0, 0, 1, 1, 1)
QR(9) = Array(1, 1, 0, 1, 0, 0, 1, 1, 0, 0, 0, 1, 1, 1, 0, 1, 1, 1, 0, 0, 1, 1, 1, 1, 1, 0, 0, 0, 1, 1, 1, 0, 0, 1, 1, 1, 1, 0, 0, 0, 1)
QR(10) = Array(1, 0, 1, 1, 1, 0, 0, 1, 0, 1, 1, 0, 1, 0, 1, 1, 1, 0, 1, 1, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 1, 0, 0, 0, 0, 1, 0, 1, 1)
QR(11) = Array(1, 1, 0, 0, 0, 0, 1, 1, 0, 1, 1, 1, 1, 0, 0, 1, 0, 1, 0, 0, 1, 0, 1, 0, 0, 0, 1, 1, 0, 1, 0, 0, 0, 1, 1, 0, 1, 0, 0, 0, 0)
QR(12) = Array(0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0, 0, 1, 0, 1, 1, 1, 1, 0, 1, 1, 0, 1, 0, 0)
QR(13) = Array(0, 0, 1, 1, 0, 1, 1, 0, 1, 1, 0, 1, 1, 1, 0, 1, 1, 0, 1, 1, 1, 0, 0, 1, 0, 0, 1, 1, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1)
QR(14) = Array(0, 1, 1, 0, 0, 0, 0, 1, 0, 1, 0, 1, 1, 1, 1, 1, 0, 0, 0, 0, 1, 1, 0, 1, 0, 1, 0, 1, 1, 1, 0, 1, 0, 0, 0, 0, 1, 0, 1, 0, 0)
QR(15) = Array(1, 0, 0, 1, 0, 1, 1, 0, 0, 0, 1, 0, 0, 0, 0, 1, 1, 0, 1, 1, 0, 1, 0, 1, 0, 1, 1, 0, 1, 1, 0, 1, 1, 1, 0, 0, 1, 1, 1, 0, 0)
QR(16) = Array(0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 1, 0, 1, 0, 1, 1, 1, 1, 1, 1, 0, 1, 0, 1, 0, 1, 0, 0, 0, 0, 1, 1, 0)
QR(17) = Array(0, 1, 0, 1, 1, 1, 1, 1, 1, 1, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 1, 1, 1, 0, 1, 0, 1, 0, 0, 0, 1, 1, 0, 1, 0, 1, 1, 0, 0, 0, 1)
QR(18) = Array(1, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 0, 0, 1, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 1, 1)
QR(19) = Array(0, 0, 0, 1, 0, 1, 1, 0, 1, 0, 1, 1, 0, 0, 0, 0, 1, 1, 0, 0, 1, 1, 1, 0, 1, 1, 0, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 1, 1)
QR(20) = Array(0, 0, 1, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 1, 0, 1, 1, 1, 0, 0, 0, 1, 0, 0, 1, 1, 1, 1, 0, 0, 0, 0, 1, 0, 1, 1, 1, 1, 0)
QR(21) = Array(0, 0, 1, 0, 1, 0, 1, 1, 0, 1, 1, 0, 1, 0, 0, 0, 0, 1, 0, 0, 1, 1, 0, 1, 0, 0, 0, 1, 1, 0, 1, 1, 0, 1, 0, 1, 1, 1, 1, 1, 0)
QR(22) = Array(1, 0, 1, 1, 0, 0, 0, 1, 0, 1, 0, 0, 0, 1, 0, 1, 1, 1, 1, 1, 0, 1, 0, 0, 0, 0, 1, 1, 0, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 0)
QR(23) = Array(0, 1, 1, 1, 1, 0, 1, 0, 1, 0, 1, 1, 0, 1, 0, 1, 1, 1, 0, 1, 1, 1, 1, 1, 0, 1, 0, 0, 0, 1, 1, 1, 1, 1, 1, 0, 1, 1, 0, 1, 0)
QR(24) = Array(1, 0, 0, 1, 1, 1, 0, 0, 1, 0, 0, 1, 1, 0, 0, 0, 1, 0, 1, 1, 0, 0, 0, 1, 0, 0, 0, 1, 1, 0, 1, 0, 1, 1, 0, 1, 0, 0, 1, 0, 1)
QR(25) = Array(1, 1, 1, 1, 0, 1, 1, 0, 1, 1, 0, 1, 0, 0, 1, 0, 0, 1, 1, 1, 0, 0, 1, 0, 1, 1, 0, 1, 1, 0, 0, 1, 1, 1, 0, 0, 0, 0, 1, 1, 1)
QR(26) = Array(1, 1, 1, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0, 1, 1, 0, 1, 1, 0, 0, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 1, 1, 0, 1)
QR(27) = Array(1, 0, 0, 1, 1, 0, 1, 0, 0, 0, 1, 1, 1, 1, 0, 1, 0, 0, 1, 1, 0, 0, 1, 0, 0, 0, 1, 1, 1, 1, 1, 0, 1, 1, 1, 1, 0, 1, 1, 0, 1)
QR(28) = Array(0, 1, 1, 0, 1, 1, 0, 0, 1, 0, 1, 0, 1, 1, 1, 0, 1, 1, 1, 1, 0, 1, 1, 1, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 1, 0, 0, 1, 1, 1, 1)
QR(29) = Array(0, 1, 1, 1, 1, 0, 1, 0, 0, 1, 0, 1, 1, 1, 1, 0, 1, 1, 1, 0, 1, 1, 0, 1, 1, 1, 1, 0, 1, 1, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0)
QR(30) = Array(0, 0, 1, 0, 1, 0, 0, 1, 1, 0, 0, 0, 1, 0, 0, 1, 1, 0, 1, 0, 1, 0, 1, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0, 1, 0, 1, 0, 1, 0, 1, 0)
QR(31) = Array(0, 1, 0, 1, 0, 0, 1, 0, 1, 1, 0, 0, 1, 0, 1, 0, 0, 0, 1, 1, 0, 0, 0, 1, 0, 1, 1, 1, 0, 1, 0, 1, 1, 1, 0, 0, 1, 1, 0, 1, 0)
QR(32) = Array(0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 1, 1, 1, 0, 1, 1, 1, 0, 0, 0, 0, 1, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1, 1)
QR(33) = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 1, 0, 0, 0, 1, 1)
QR(34) = Array(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, 0, 1, 1, 0, 1, 0, 1, 1, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 1, 0, 1, 0, 1, 0, 0, 1, 1, 1)
QR(35) = Array(0, 1, 1, 1, 1, 1, 0, 1, 1, 1, 0, 1, 1, 1, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 1, 1, 1, 0, 1, 0, 0, 0)
QR(36) = Array(0, 1, 0, 0, 0, 1, 0, 1, 1, 1, 1, 0, 1, 1, 1, 0, 0, 1, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0)
QR(37) = Array(0, 1, 0, 0, 0, 1, 0, 1, 0, 0, 1, 1, 0, 1, 0, 1, 0, 0, 1, 0, 1, 0, 1, 1, 1, 0, 0, 1, 0, 1, 0, 0, 1, 0, 1, 1, 1, 1, 0, 1, 0)
QR(38) = Array(0, 1, 0, 0, 0, 1, 0, 1, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 1, 1, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0, 1, 1, 1, 0, 0, 0, 1, 1, 1, 0)
QR(39) = Array(0, 1, 1, 1, 1, 1, 0, 1, 1, 1, 1, 0, 1, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0, 1, 0, 1, 0, 1, 0, 1, 0, 0, 0, 0, 1, 0, 0, 1, 1)
QR(40) = Array(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 1, 0, 1, 0, 1, 1, 1, 0, 1, 1, 0, 1, 0, 1, 1, 0, 0, 0, 1, 1, 0, 1, 0, 0, 1, 0, 0, 0, 1)
DW = QRUp.Width / 42
With QRUp
    .Width = .Width / 42 * 41
End With
With QRDo
    .Width = .Width / 42 * 41
    .Height = .Height / 21 * 20
End With
Dim KK As Integer
Dim cL As Long
For k = 0 To 40
    For KK = 0 To 40
        If QR(KK)(k) = 0 Then
            cL = 0
        Else
            cL = vbWhite
        End If
        QRUp.ForeColor = cL
        QRDo.ForeColor = cL
        QRUp.FillColor = cL
        QRDo.FillColor = cL
        If k <= 20 Then
            QRUp.Line (KK * DW, k * DW)-Step(DW, DW), , B
        Else
            QRDo.Line (KK * DW, (k - 21) * DW)-Step(DW, DW), , B
        End If
    Next KK
Next k
Reset
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
If Not nM Then
    '点红叉叉
    Shell App.Path & "\" & App.EXEName, vbNormalFocus
    End
Else
    Unload Lv14
End If
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
                Dim bb As Integer, BBB As Boolean
                bb = Qiang(, B + 1) - 9
                If bb >= 0 And bb <= 3 And W(3).Visible Then
                    WH(bb) = (Val(WH(bb)) + 1) Mod 10
                    BBB = True
                    For bb = 0 To 3
                        If PassWd(bb) <> Val(WH(bb)) Then
                            BBB = False
                        End If
                    Next bb
                    If BBB Then
                        Wall(6).Exist = False
                        Wall(4).Exist = False
                        W(5).Visible = False
                        W(3).Visible = False
                    End If
                End If
            End If
        Else
            If Qiang(, B) Or Grnd Then
                .Speed_Ver = 0
                Do While Qiang(, B)
                    B = B + 1
                Loop
                '踩
                Select Case Qiang(, B - 1)
                    Case 8 'Win
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
W(5).Visible = True
W(3).Visible = True
Wall(6).Exist = True
Wall(3).Exist = False
Wall(4).Exist = True
Wall(5).Exist = False
Dim k As Integer
For k = 0 To 3
    PassWd(k) = Int(Rnd() * 10)
Next k
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
'画上正确密码
Dim NumIm(0 To 9) As String
NumIm(0) = "111101101101111"
NumIm(1) = "010110010010111"
NumIm(2) = "111001111100111"
NumIm(3) = "111001111001111"
NumIm(4) = "101101111001001"
NumIm(5) = "111100111001111"
NumIm(6) = "111100111101111"
NumIm(7) = "111001001001001"
NumIm(8) = "111101111101111"
NumIm(9) = "111101111001111"
k = 0
Dim ZongX As Integer
Dim GbX As Integer, GbY As Integer
Dim cL As Long
For ZongX = 13 To 25 Step 4
    For GbY = 18 To 22
        For GbX = ZongX To IIf(k = 3, ZongX + 2, ZongX + 3)
            If GbX = ZongX + 3 Then
                cL = vbWhite
            Else
                If Mid(NumIm(PassWd(k)), (GbY - 18) * 3 + GbX - ZongX + 1, 1) = "1" Then
                    cL = 0
                Else
                    cL = vbWhite
                End If
            End If
            QRUp.ForeColor = cL
            QRDo.ForeColor = cL
            QRUp.FillColor = cL
            QRDo.FillColor = cL
            If GbY <= 19 Then
                QRDo.Line (GbX * DW, GbY * DW)-Step(DW, DW), , B
            Else
                QRUp.Line (GbX * DW, (GbY - 20) * DW)-Step(DW, DW), , B
            End If
        Next GbX
    Next GbY
    k = k + 1
Next ZongX
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
With Source
    .Move X - DragX, Y - DragY
    If Abs(QRUp.Left - QRDo.Left) <= 106 Then
        QRDo.Left = QRUp.Left
    End If
    If Abs(QRUp.Top + QRUp.Height - QRDo.Top) <= 106 Then
        QRDo.Top = QRUp.Top + QRUp.Height
    End If
    If Abs(QRDo.Top + QRDo.Height - QRUp.Top) <= 106 Then
        QRUp.Top = QRDo.Top + QRDo.Height
    End If
End With
End Sub

Private Sub QRDo_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub QRDo_KeyUp(KeyCode As Integer, Shift As Integer)
Form_KeyUp KeyCode, Shift
End Sub

Private Sub QRUp_DragDrop(Source As Control, X As Single, Y As Single)
Form_DragDrop Source, X + QRUp.Left, Y + QRUp.Top
End Sub

Private Sub QRDo_DragDrop(Source As Control, X As Single, Y As Single)
Form_DragDrop Source, X + QRDo.Left, Y + QRDo.Top
End Sub

Private Sub QRUp_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub QRUp_KeyUp(KeyCode As Integer, Shift As Integer)
Form_KeyUp KeyCode, Shift
End Sub

Private Sub QRUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With QRUp
    .Drag 1
End With
DragX = X
DragY = Y
End Sub

Private Sub QRDo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With QRDo
    .Drag 1
End With
DragX = X
DragY = Y
End Sub

Private Sub Quit_Click()
Unload Me
End Sub

Private Sub Suicide_Click()
Si
End Sub

Private Sub Win()
If Val(GetSetting("iGlope", "woqu", "lv", 1)) < 18 Then
    SaveSetting "iGlope", "woqu", "lv", "18"
End If
MsgBox "恭喜你打通了" & Me.Name & "！按OK进入下一关。", vbOKOnly + vbInformation, Me.Caption
Lv15.Visible = True
nM = True
Unload Me
End Sub

Private Sub WH_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
Form_DragDrop Source, X + WH(Index).Left, Y + WH(Index).Top
End Sub


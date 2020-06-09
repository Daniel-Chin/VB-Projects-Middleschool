VERSION 5.00
Begin VB.Form MainMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "我去冒险"
   ClientHeight    =   5652
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9600
   DrawWidth       =   3
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   22.2
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5652
   ScaleWidth      =   9600
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer EcY 
      Interval        =   36
      Left            =   7800
      Top             =   2880
   End
   Begin VB.Timer Yser 
      Interval        =   36
      Left            =   8160
      Top             =   4800
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "<< 返回"
      ForeColor       =   &H00000000&
      Height          =   504
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   4366
      Visible         =   0   'False
      Width           =   9672
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "奖励关 >>"
      ForeColor       =   &H00000000&
      Height          =   504
      Index           =   6
      Left            =   0
      TabIndex        =   9
      Top             =   3866
      Visible         =   0   'False
      Width           =   9672
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "基本关卡 >>"
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   5
      Left            =   0
      TabIndex        =   8
      Top             =   3366
      Visible         =   0   'False
      Width           =   9672
   End
   Begin VB.Label L 
      BackColor       =   &H00FF0000&
      Height          =   132
      Index           =   3
      Left            =   600
      TabIndex        =   7
      Top             =   2760
      Width           =   1212
   End
   Begin VB.Label L 
      BackColor       =   &H00FF0000&
      Height          =   132
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   480
      Width           =   1212
   End
   Begin VB.Label L 
      BackColor       =   &H00FF0000&
      Height          =   2412
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   480
      Width           =   132
   End
   Begin VB.Label L 
      BackColor       =   &H00FF0000&
      Height          =   2412
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   132
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "检查并更新"
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   4
      Left            =   0
      TabIndex        =   3
      Top             =   4866
      Width           =   9672
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "使用秘籍"
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   3
      Left            =   0
      TabIndex        =   2
      Top             =   4366
      Width           =   9672
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "神器库"
      ForeColor       =   &H00000000&
      Height          =   504
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   3866
      Width           =   9672
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "开始游戏  >>"
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   3366
      Width           =   9666
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nM As Boolean
Dim GB As Integer
Dim Ms As Byte
Dim cx As Long, cy As Long
Dim EcX As Double
Const qR As Integer = 180
Dim qX(0 To 1) As Long, qY(0 To 1) As Long
Const Pi As Double = 3.14159265358979

Private Sub Bt_Click(Index As Integer)
Dim k As Byte
Select Case Index
    Case 0
        For k = 0 To 6
            Bt(k).Visible = Not Bt(k).Visible
        Next k
        GB = 5
        Renew
    Case 1
        '神器
        SQ.Show
        nM = True
        Unload Me
    Case 3
        '秘籍
        MiJi.Show
        nM = True
        Unload Me
    Case 4
        '更新
        GX1.Show
        nM = True
        Unload Me
    Case 5
        ChLv.Show
        nM = True
        Unload Me
    Case 6
        If GetSetting("iGlope", "woqu", "lv", 1) = 18 Then
            ChS.Show
            nM = True
            Unload Me
        Else
            MsgBox "奖励关卡还没有解锁。请先通过全部基本关卡。", vbExclamation + vbOKOnly
        End If
    Case 2
        For k = 0 To 6
            Bt(k).Visible = Not Bt(k).Visible
        Next k
        GB = 0
        Renew
End Select
End Sub

Private Sub Bt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If GB <> Index Then
    GB = Index
    Renew
End If
Form_MouseMove Button, Shift, X + Bt(Index).Left, Y + Bt(Index).Top
End Sub

Private Sub EcY_Timer()
EcX = MODx((EcX + 0.1))
Dim YS As Long, cu As Long
If EcX >= Pi Then
    YS = vbWhite
    cu = 66
End If
Dim k As Byte
For k = 0 To 1
    Line (qX(k) - Cos(EcX) * qR + cu, qY(k) - Sin(EcX) * qR + cu)-(qX(k) + Cos(EcX) * qR - cu, qY(k) + Sin(EcX) * qR - cu), YS
Next k
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim pS(), k
pS = Array(1, 3, 5, 4, 0, 6, 2)
Select Case KeyCode
    Case 38 'Up
        For k = 0 To 6
            If pS(k) = GB Then
                GB = k
                Exit For
            End If
        Next k
    Case 40 'Down
        GB = pS(GB)
    Case 13
        Bt_Click GB
        Exit Sub
    Case 27
        If Bt(2).Visible Then
            For k = 0 To 6
                Bt(k).Visible = Not Bt(k).Visible
            Next k
            GB = 0
        Else
            Unload Me
        End If
    Case Else
        Exit Sub
End Select
Renew
End Sub

Private Sub Form_Load()
''''''''''''''''
'Shop.Show
'Unload Me
'Exit Sub
''''''''''''''''
If Val(Command) >= 1 Then
    SQMsgBox.Show
    SQMsgBox.SQ_ID = Val(Command)
    nM = True
    Unload Me
    Exit Sub
End If
Music.Visible = False
Music.ModeSwitch "menu"
Show
PaintPicture LoadPicture(App.Path & "\logo.jpg"), 3000, 436, 3666, 2666
Dim k As Integer
For k = 0 To 3
    L(k).ToolTipText = "这只是一个没有意义的墨水池，而已。"
Next k
GB = 0
Ms = 255
Renew
qX(0) = qR + 66
qY(0) = qR + 66
qX(1) = 9672 - qR - 166
qY(1) = qR + 66
End Sub

Private Sub Renew()
With Bt(GB)
    .BackColor = 0
    .ForeColor = vbWhite
End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Ms < 255 Then
    Line (X, Y)-(cx, cy), RGB(Ms, Ms, 255)
    cx = X
    cy = Y
    Ms = Ms + 17
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not nM Then
    End
End If
End Sub

Private Sub L_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Ms = 0
cx = X + L(Index).Left
cy = Y + L(Index).Top
End Sub

Private Sub Yser_Timer()
Dim k As Byte
For k = 0 To 6
    With Bt(k)
        If .ForeColor <> 0 And k <> GB Then
            .BackColor = .BackColor + RGB(51, 51, 51)
            .ForeColor = .ForeColor - RGB(51, 51, 51)
        End If
    End With
Next k
End Sub

Function MODx(A As Double, Optional B As Double = 2 * Pi) As Double
If B > 0 Then
    Do Until A >= 0 And A < B
        If A < 0 Then
            A = A + B
        Else
            A = A - B
        End If
    Loop
    MODx = A
Else
    Error 17
End If
End Function


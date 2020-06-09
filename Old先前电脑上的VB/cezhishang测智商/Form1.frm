VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "测智商"
   ClientHeight    =   6828
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10068
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6828
   ScaleWidth      =   10068
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Eer 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   8280
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton Ex 
      Caption         =   "Quit Game"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   8040
      TabIndex        =   3
      Top             =   600
      Width           =   1812
   End
   Begin VB.CommandButton P 
      Caption         =   "WASD to move"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   3828
      TabIndex        =   2
      Top             =   2760
      Width           =   2412
   End
   Begin VB.Timer Timer1 
      Left            =   8400
      Top             =   2520
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   1440
      X2              =   1200
      Y1              =   1560
      Y2              =   1320
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   1200
      X2              =   1080
      Y1              =   1320
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   1200
      X2              =   1320
      Y1              =   1320
      Y2              =   1920
   End
   Begin VB.Label Label3 
      Caption         =   "注意了，通关所用的时间决定了你的成绩！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   2172
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "?"
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
      Height          =   852
      Left            =   8040
      TabIndex        =   4
      Top             =   5400
      Width           =   612
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   120
      X2              =   9840
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   252
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   252
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   252
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   372
      Left            =   1320
      Top             =   4080
      Width           =   372
   End
   Begin VB.Label zS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "智商：250"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1812
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Level - 1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   2208
      TabIndex        =   0
      Top             =   360
      Width           =   5652
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZhiShang As Integer

Private Sub Eer_Click()
End
End Sub

Private Sub Ex_Click()
Ex.Visible = False
Label1.Visible = False
P.Visible = False
zS.Visible = False
Label3.Visible = False
Eer.Visible = True
Me.FontSize = 40
Show
Print "你赢了！"
Print "玩这个游戏就是"
Print "在比谁能更早意"
Print "识到，这个游戏"
Print "除了退出以外没"
Print "有任何其他功能！"
Print "您的智商：" & ZhiShang
Timer1.Interval = 0
End Sub

Private Sub Form_Load()
MsgBox "点击确定开始测智商。", vbOKOnly + vbQuestion, "测智商"
Timer1.Interval = 300
ZhiShang = 250
End Sub

Private Sub P_KeyPress(KeyAscii As Integer)
With P
    Select Case KeyAscii
        Case 119
            .Top = .Top - 40
        Case 115
            .Top = .Top + 40
        Case 97
            .Left = .Left - 40
        Case 100
            .Left = .Left + 40
    End Select
End With
End Sub

Private Sub P_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.BackColor = vbBlack
End Sub

Private Sub P_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.BackColor = RGB(230, 230, 230)
End Sub

Private Sub Timer1_Timer()
ZhiShang = ZhiShang - 1
zS = "智商：" & ZhiShang
End Sub

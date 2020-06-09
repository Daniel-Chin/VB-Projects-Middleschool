VERSION 5.00
Begin VB.Form Fake 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "奇怪的小冒险"
   ClientHeight    =   6276
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10872
   Icon            =   "Fake.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6276
   ScaleWidth      =   10872
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Start 
      BackColor       =   &H00FFFFFF&
      Caption         =   "　商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Index           =   2
      Left            =   6480
      TabIndex        =   3
      Top             =   4680
      Width           =   3612
   End
   Begin VB.Label Start 
      BackColor       =   &H00FFFFFF&
      Caption         =   "　商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   3840
      Width           =   3612
   End
   Begin VB.Label Start 
      BackColor       =   &H00FFFFFF&
      Caption         =   "　商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   2520
      Width           =   3612
   End
   Begin VB.Label BiaoTi 
      Caption         =   "奇怪的小冒险V1.7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   1296
      TabIndex        =   0
      Top             =   600
      Width           =   8292
   End
End
Attribute VB_Name = "Fake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nm As Boolean
Dim NowOn As Byte
Dim mode As Byte

Private Sub Form_Load()
Nm = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        For I = 0 To 2
            With Start(I)
                If I = NowOn Then GoTo Skip
                .BackColor = vbWhite
                .ForeColor = 0
            End With
Skip:   Next I
        NowOn = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not Nm Then
    End
End If
End Sub

Private Sub Start_Click(Index As Integer)
If InputBox("这个游戏根本没有隐藏彩蛋！放弃吧！（使用ESC退出）", "There's no Ester Eggs here! Turn back!", "输入错误代码的倒序！") = "U_ARE_N_idiot" Then
    Nm = True
    Shop.Visible = True
    Unload Me
Else
    MsgBox "可惜你没有把它记下来。", , ""
    End
End If
End Sub

Private Sub Start_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove 0, 0, 0, 0
With Start(Index)
    .BackColor = 0
    .ForeColor = vbWhite
End With
NowOn = Index
End Sub


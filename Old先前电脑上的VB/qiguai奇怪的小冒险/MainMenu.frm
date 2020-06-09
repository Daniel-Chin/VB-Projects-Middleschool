VERSION 5.00
Begin VB.Form MainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "奇怪的小冒险"
   ClientHeight    =   6276
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10872
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6276
   ScaleWidth      =   10872
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tM 
      Left            =   7080
      Top             =   3480
   End
   Begin VB.Label Corner 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   612
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   732
   End
   Begin VB.Label Start 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "更新游戏"
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
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "设置"
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
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "开始游戏"
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
      Alignment       =   2  'Center
      Caption         =   "奇怪的小冒险"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1812
      Left            =   990
      TabIndex        =   0
      Top             =   480
      Width           =   8892
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nm As Boolean
Dim NowOn As Byte
Dim mode As Byte

Private Sub EsterEgg()
Nm = True
Egg1.Visible = True
Unload Me
End Sub

Private Sub Corner_Click()
Call EsterEgg
End Sub

Private Sub Corner_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
With Corner
    .Caption = "爽"
    .BackColor = 0
End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 101 'e
        Nm = True
        Egg2.Visible = True
        Unload Me
End Select
End Sub

Private Sub Form_Load()
Nm = False
If Theme.TongGuan Then
    With BiaoTi
        .Caption = "你以为这个游戏只有这无聊的三关，而且第三关还通不过？"
        .FontSize = 20
    End With
    Tm.Interval = 3000
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
        For I = 0 To 2
            With Start(I)
                If I = NowOn Then GoTo Skip
                .BackColor = vbWhite
                .ForeColor = 0
            End With
Skip:   Next I
        NowOn = 3
        With Corner
            .Caption = ""
            .BackColor = Me.BackColor
        End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not Nm Then
    End
End If
End Sub

Private Sub Start_Click(Index As Integer)
Select Case Index
    Case 0  '开始
        Level1.Visible = True
        Nm = True
        Unload Me
    Case 1  '设置
        Settings.Visible = True
        Nm = True
        Unload Me
    Case 2  '更新
        RenewAsk.Visible = True
        Nm = True
        Unload Me
    Case Else
        MsgBox "This message will never appear on your screen. Isn't that weird?", vbCritical + vbYesNo, "One of the Ester Eggs"
End Select
End Sub

Private Sub Start_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
Form_MouseMove 0, 0, 0, 0
With Start(Index)
    .BackColor = 0
    .ForeColor = vbWhite
End With
NowOn = Index
End Sub

Private Sub Tm_Timer()
mode = mode + 1
Select Case mode
    Case 1
        BiaoTi = BiaoTi & "这怎么可能呢？"
    Case 2
        BiaoTi = "这个游戏的主体其实是隐藏的，你必须将它们找到，才能继续这个风骚的游戏！"
    Case 3
        BiaoTi = BiaoTi & "提示：左上角"
        Tm.Interval = 0
        mode = 0
        Theme.TongGuan = False
End Select
End Sub

Sub FUCK()
Nm = True
Egg99.Visible = True
Unload Me
End Sub


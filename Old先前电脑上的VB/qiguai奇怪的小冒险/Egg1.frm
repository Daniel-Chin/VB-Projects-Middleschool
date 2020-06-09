VERSION 5.00
Begin VB.Form Egg1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "真正奇怪的冒险"
   ClientHeight    =   6672
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10992
   Icon            =   "Egg1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6672
   ScaleWidth      =   10992
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Key 
      Height          =   372
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "另一个秘密"
      Top             =   -500
      Width           =   972
   End
   Begin VB.Timer Main 
      Interval        =   40
      Left            =   9600
      Top             =   3480
   End
   Begin VB.Timer Kill 
      Left            =   8520
      Top             =   3360
   End
   Begin VB.CommandButton Start 
      Caption         =   "开始喽"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   3810
      TabIndex        =   0
      Top             =   2730
      Width           =   3372
   End
   Begin VB.Label Win 
      BackColor       =   &H00C00000&
      Height          =   492
      Left            =   9840
      TabIndex        =   19
      Top             =   0
      Width           =   492
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   1932
      Index           =   17
      Left            =   1680
      TabIndex        =   18
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   16
      Left            =   480
      TabIndex        =   17
      Top             =   1440
      Width           =   2532
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   492
      Index           =   15
      Left            =   1200
      TabIndex        =   16
      Top             =   2160
      Width           =   2412
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   252
      Index           =   14
      Left            =   1200
      TabIndex        =   15
      Top             =   3240
      Width           =   1692
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   1932
      Index           =   13
      Left            =   1080
      TabIndex        =   14
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   1932
      Index           =   12
      Left            =   8400
      TabIndex        =   13
      Top             =   360
      Width           =   252
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   252
      Index           =   11
      Left            =   600
      TabIndex        =   12
      Top             =   5040
      Width           =   7692
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   252
      Index           =   10
      Left            =   1560
      TabIndex        =   11
      Top             =   5520
      Width           =   7812
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   1932
      Index           =   9
      Left            =   9120
      TabIndex        =   10
      Top             =   3600
      Width           =   252
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   8
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   9252
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   6132
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   492
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   6120
      Width           =   10332
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   6132
      Index           =   5
      Left            =   10440
      TabIndex        =   6
      Top             =   240
      Width           =   252
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   4
      Left            =   7920
      TabIndex        =   5
      Top             =   2640
      Width           =   2652
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   1932
      Index           =   3
      Left            =   7680
      TabIndex        =   4
      Top             =   1560
      Width           =   252
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   2
      Left            =   3360
      TabIndex        =   3
      Top             =   2160
      Width           =   4332
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   732
      Index           =   1
      Left            =   2280
      TabIndex        =   2
      Top             =   4080
      Width           =   6972
   End
   Begin VB.Label Wall 
      BackColor       =   &H00000000&
      Height          =   1932
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   2400
      Width           =   252
   End
End
Attribute VB_Name = "Egg1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Started As Boolean
Dim xx As Long
Dim LE As Long
Dim Nm As Boolean
Dim z As Boolean, s As Boolean, y As Boolean, X As Boolean

Private Sub Form_Load()
s = False
X = False
z = False
y = False
Start.Caption = "开始"
LE = Wall(12).Left
Nm = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Started Then
    If y < 2000 Then
        Kill.Interval = 50
    End If
    xx = X
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not Nm Then
    End
End If
End Sub

Private Sub Key_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37     '左
        z = True
    Case 38     '上
        s = True
    Case 39    '右
        y = True
    Case 40
        X = True '下
End Select
End Sub

Private Sub Key_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37     '左
        z = False
    Case 38     '上
        s = False
    Case 39    '右
        y = False
    Case 40
        X = False '下
End Select
End Sub

Private Sub Kill_Timer()
With Wall(12)
    .Left = .Left - .Width / 2
    If xx > .Left And xx < 6500 Then
        Die
    End If
End With
End Sub

Private Sub Die()
s = False
X = False
z = False
y = False
Started = False
Start.Visible = True
Start.Caption = "重试"
MsgBox "你死了。", vbCritical + vbOKOnly, "遗憾"
If Rnd() < 0.1 Then
    MsgBox "如果你忘记了游戏操作方法，请参见方块游戏第一关！", vbCritical + vbOKOnly, "你忘了应该用方向键控制么？？？"
End If
Wall(12).Left = LE
Kill.Interval = 0
For I = 0 To 17
    Wall(I).Visible = False
Next I
End Sub

Private Sub Main_Timer()
With Start
    If s Then
        .Top = .Top - 40
    End If
    If z Then
        .Left = .Left - 40
    End If
    If y Then
        .Left = .Left + 40
    End If
    If X Then
        .Top = .Top + 40
    End If
End With
End Sub

Private Sub Start_Click()
Started = True
Start.Visible = False
For I = 0 To 17
    Wall(I).Visible = True
Next I
End Sub

Private Sub Wall_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
If Started Then
    Die
End If
End Sub

Private Sub Win_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Nm = True
MsgBox "你赢了。", vbOKOnly + vbInformation, "恭喜"
Egg2.Visible = True
Unload Me
End Sub

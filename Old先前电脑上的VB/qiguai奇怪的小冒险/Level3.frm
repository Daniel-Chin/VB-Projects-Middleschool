VERSION 5.00
Begin VB.Form Level3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "奇怪的小冒险 - 第3关"
   ClientHeight    =   5424
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10176
   Icon            =   "Level3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5424
   ScaleWidth      =   10176
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Back 
      Caption         =   "回到主菜单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   6840
      TabIndex        =   2
      Top             =   4200
      Width           =   3012
   End
   Begin VB.Timer Main 
      Interval        =   30
      Left            =   8160
      Top             =   720
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1332
      Left            =   5880
      Top             =   2400
      Width           =   4212
   End
   Begin VB.Label Win 
      BackColor       =   &H0000FFFF&
      Caption         =   "★"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   852
      Left            =   8880
      TabIndex        =   1
      Top             =   1560
      Width           =   852
   End
   Begin VB.Shape I 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   612
      Left            =   1320
      Top             =   1080
      Width           =   852
   End
   Begin VB.Label Ground 
      BackColor       =   &H00000000&
      Caption         =   "游戏变难了：你从现在起不能跳。"
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
      Height          =   1692
      Left            =   600
      TabIndex        =   0
      Top             =   2400
      Width           =   3252
   End
End
Attribute VB_Name = "Level3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Speed As Byte = 100
Dim VerSpeed As Integer
Dim Nm As Boolean
Dim z As Boolean, Y As Boolean

Private Sub Back_Click()
Nm = True
MainMenu.Visible = True
Unload Me
End Sub

Private Sub Back_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub Back_KeyUp(KeyCode As Integer, Shift As Integer)
Form_KeyUp KeyCode, Shift
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37     '左
        z = True
    Case 39    '右
        Y = True
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37     '左
        z = False
    Case 39    '右
        Y = False
End Select
End Sub

Private Sub Form_Load()
Nm = False
If Theme.TongGuan Then
    Back.Visible = True
    Ground.Caption = "你不能移动！放弃吧。"
Else
    Back.Visible = False
    Ground.Caption = "游戏变难了：你从现在起不能跳。"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not Nm Then
    End
End If
End Sub

Private Sub Main_Timer()
I.Top = I.Top - VerSpeed
If I.Top + I.Height < Ground.Top Then
    VerSpeed = VerSpeed - 5
Else
    VerSpeed = 0
End If
If I.Left + I.Width > Win.Left And I.Left < Win.Left + Win.Width And I.Top < Win.Top + Win.Height And I.Top + I.Height > Win.Top Then
    MsgBox "现在！！！回到第1关！", vbOKOnly + vbInformation, "恭喜"
    Level1.Visible = True
    Theme.TongGuan = True
    Nm = True
    z = False
    Y = False
    Unload Me
    Exit Sub
End If

    If z Then
        If I.Left > 0 Then
            I.Left = I.Left - Speed
        End If
    End If
    If Y Then
        If I.Left + I.Width < Me.Width Then
            I.Left = I.Left + Speed
        End If
    End If
End Sub



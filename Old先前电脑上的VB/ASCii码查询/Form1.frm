VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "ASCii"
   ClientHeight    =   4488
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7440
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4488
   ScaleWidth      =   7440
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   3240
      Top             =   2040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "段落查询"
      Height          =   492
      Left            =   4200
      TabIndex        =   7
      Top             =   3720
      Width           =   1932
   End
   Begin VB.CommandButton Command2 
      Caption         =   "生僻字符查询"
      Height          =   492
      Left            =   720
      TabIndex        =   6
      Top             =   3720
      Width           =   1932
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   4200
      TabIndex        =   5
      Text            =   "65"
      Top             =   2160
      Width           =   2772
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   60
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   960
      TabIndex        =   4
      Text            =   "A"
      Top             =   2160
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "最大化"
      Height          =   252
      Left            =   6480
      TabIndex        =   2
      Top             =   0
      Width           =   732
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   4
      Index           =   4
      X1              =   2520
      X2              =   2280
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   4
      Index           =   3
      X1              =   3960
      X2              =   3720
      Y1              =   2640
      Y2              =   2400
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   4
      Index           =   2
      X1              =   3960
      X2              =   3720
      Y1              =   2640
      Y2              =   2880
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   4
      Index           =   1
      X1              =   2520
      X2              =   2280
      Y1              =   2400
      Y2              =   2640
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   4
      Index           =   0
      X1              =   2280
      X2              =   3960
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Caption         =   "码查询"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   972
      Left            =   4200
      TabIndex        =   3
      Top             =   840
      Width           =   3012
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   3960
      X2              =   3960
      Y1              =   1080
      Y2              =   960
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   3720
      X2              =   3720
      Y1              =   1080
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   5
      X1              =   3960
      X2              =   3960
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   4
      X1              =   3720
      X2              =   3720
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   3360
      X2              =   2760
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   2760
      X2              =   2760
      Y1              =   840
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   3360
      X2              =   2760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2280
      X2              =   2040
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2280
      X2              =   2400
      Y1              =   1320
      Y2              =   1560
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1920
      X2              =   2040
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2400
      X2              =   1920
      Y1              =   1560
      Y2              =   1680
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   2400
      X2              =   1920
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1440
      X2              =   960
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1200
      X2              =   1560
      Y1              =   840
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1200
      X2              =   840
      Y1              =   840
      Y2              =   1680
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "  X"
      BeginProperty Font 
         Name            =   "ADMUI3Sm"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   7200
      TabIndex        =   1
      Top             =   0
      Width           =   252
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim position As POINTAPI
Dim Draged As Boolean
Dim x As Long, y As Long
Private Sub Command1_Click()
Me.Top = 0
Me.Left = Screen.Width - Me.Width
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.BackColor = &HFF&
End Sub
Private Sub Command2_Click()
Form2.Visible = True
End Sub
Private Sub Command3_Click()
Form3.Visible = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.BackColor = &HFF&
End Sub
Private Sub Label1_Click()
Draged = False
Timer1.Interval = 0
End Sub
Private Sub Label1_DblClick()
If Draged = False Then
    Timer1.Interval = 40
    Draged = True
    GetCursorPos position
    x = position.x
    y = position.y
End If
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.BackColor = &HFF&
End Sub
Private Sub Label2_Click()
Unload Form1
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.BackColor = &HC0C0FF
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Text1 = ""
Text2 = KeyAscii
If Text2 = "22" Then
    MsgBox "若想使用复制粘贴，请使用“生僻字符查询”。若想查询多个字符，请使用“段落查询。”", vbCritical, "错误"
    Text2 = ""
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1 = Chr(Val(Text2))
End If
End Sub
Private Sub Timer1_Timer()
GetCursorPos position
Me.Left = Me.Left + (position.x - x) * 12
Me.Top = Me.Top + (position.y - y) * 12
x = position.x
y = position.y
End Sub

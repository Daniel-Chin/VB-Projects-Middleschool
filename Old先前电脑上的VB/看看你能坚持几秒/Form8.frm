VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   Caption         =   "本游戏由麻瓜传记出品"
   ClientHeight    =   6012
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11232
   LinkTopic       =   "Form8"
   ScaleHeight     =   6012
   ScaleWidth      =   11232
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   9240
      Top             =   2280
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   1200
      TabIndex        =   2
      Text            =   "喵！"
      Top             =   0
      Width           =   972
   End
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   9480
      Top             =   1680
   End
   Begin VB.Image Image1 
      Height          =   2436
      Left            =   3840
      Picture         =   "Form8.frx":0000
      Top             =   3120
      Width           =   3396
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The Tale of Two Maguas"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   22.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   612
      Left            =   3000
      TabIndex        =   1
      Top             =   2160
      Width           =   5052
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "麻瓜传记"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1692
      Left            =   4560
      TabIndex        =   0
      Top             =   360
      Width           =   1812
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Private Sub Form_Load()
Form1.x_s = 12
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
a = a & Chr(KeyAscii)
If a = "kip" Then
    Call Timer1_Timer
End If
End Sub
Private Sub Timer1_Timer()
Unload Form8
Form7.Visible = True
End Sub
Private Sub Timer2_Timer()
Text1.SetFocus
Timer2.Interval = 0
End Sub

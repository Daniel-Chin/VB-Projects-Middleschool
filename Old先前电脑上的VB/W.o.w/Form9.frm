VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   0  'None
   Caption         =   "真相大白！！！！！"
   ClientHeight    =   4884
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5988
   LinkTopic       =   "Form9"
   ScaleHeight     =   4884
   ScaleWidth      =   5988
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   5772
   End
   Begin VB.Label Label1 
      Caption         =   $"Form9.frx":0000
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4692
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   4896
      Left            =   0
      Picture         =   "Form9.frx":01F1
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nmal As Boolean
Private Sub Form_Unload(Cancel As Integer)
If nmal = True Then
Else
    End
End If
End Sub
Private Sub Command1_Click()
nmal = True
End
End Sub

Private Sub Form_Load()
Form11.WindowsMediaPlayer1.URL = App.Path & "\.W.o.w_files\mellon.wma"
Form11.WindowsMediaPlayer1.Controls.currentPosition = 1
Form11.WindowsMediaPlayer1.Controls.play
Form11.a_b = True
nmal = False
Form8.lala_la = True
End Sub

Private Sub Image1_Click()
Label1.Visible = True
Command1.Visible = True
End Sub

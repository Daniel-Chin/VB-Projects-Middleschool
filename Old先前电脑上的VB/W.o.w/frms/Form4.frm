VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000008&
   ClientHeight    =   8676
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   12276
   LinkTopic       =   "Form4"
   ScaleHeight     =   8676
   ScaleWidth      =   12276
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer4 
      Left            =   6000
      Top             =   3000
   End
   Begin VB.Timer Timer2 
      Interval        =   4000
      Left            =   7800
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7080
      Top             =   720
   End
   Begin VB.Image Image5 
      Height          =   9216
      Left            =   0
      Picture         =   "Form4.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   12288
   End
   Begin VB.Image Image4 
      Height          =   9216
      Left            =   0
      Picture         =   "Form4.frx":251CF
      Top             =   0
      Visible         =   0   'False
      Width           =   12288
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   9216
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12492
   End
   Begin VB.Image Image3 
      Height          =   9216
      Left            =   0
      Picture         =   "Form4.frx":379AF
      Top             =   0
      Visible         =   0   'False
      Width           =   12288
   End
   Begin VB.Image Image2 
      Height          =   9216
      Left            =   0
      Picture         =   "Form4.frx":6F2F3
      Top             =   0
      Visible         =   0   'False
      Width           =   12288
   End
   Begin VB.Image Image1 
      Height          =   9216
      Left            =   0
      Picture         =   "Form4.frx":A7431
      Top             =   0
      Visible         =   0   'False
      Width           =   12288
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time As Long
Dim time2 As Long
Dim time3 As Long
Dim nmal As Boolean
Private Sub Form_Unload(Cancel As Integer)
If nmal = True Then
Else
    End
End If
End Sub
Private Sub 结束()
nmal = True
Unload Form4
Form10.Visible = True
End Sub
Private Sub Form_Load()
time = -2
nmal = False
Form11.WindowsMediaPlayer1.Controls.currentPosition = 0
Label1 = "大家好，我是" & Form3.player & "。不，我曾经是。现在我只是一个代号，" & Form3.playercd & "。"
End Sub

Private Sub Timer1_Timer()
time = time + 1
If time = 2 Then Image4.Visible = True: Label1.Visible = False
If time = 3 Then Image4.Visible = False: Image5.Visible = True
If time = 4 Then Image5.Visible = False
If time = 12 Then Timer4.Interval = 500
If time = 14 Then 结束
End Sub
Private Sub Timer2_Timer()
time2 = time2 + 1
If time2 Mod 3 = 0 Then
    Image1.Visible = True
    Image2.Visible = False
    Image3.Visible = False
End If
If time2 Mod 3 = 1 Then
    Image1.Visible = False
    Image2.Visible = True
    Image3.Visible = False
End If
If time2 Mod 3 = 2 Then
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = True
End If
End Sub
Private Sub Timer4_Timer()
Form11.WindowsMediaPlayer1.settings.volume = Form11.WindowsMediaPlayer1.settings.volume - 12
End Sub

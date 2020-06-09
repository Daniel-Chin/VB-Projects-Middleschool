VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "关卡加载"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14964
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   7680
   ScaleWidth      =   14964
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6840
      Top             =   3600
   End
   Begin VB.Image Image2 
      Height          =   384
      Left            =   1440
      Picture         =   "Form6.frx":38CDE
      Top             =   7200
      Width           =   12012
   End
   Begin VB.Image Image1 
      Height          =   7668
      Left            =   0
      Picture         =   "Form6.frx":3CAEB
      Top             =   0
      Width           =   15000
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Dim nmal As Boolean
Private Sub Form_Unload(Cancel As Integer)
If nmal = True Then
Else
    End
End If
End Sub

Private Sub Form_Load()
Form11.WindowsMediaPlayer1.URL = App.Path & "\.W.o.w_files\ahahah.mp3"
Form11.WindowsMediaPlayer1.Controls.currentPosition = 0
Form11.WindowsMediaPlayer1.Controls.play
Image2.Width = 15
nmal = False
End Sub
Private Sub Timer1_Timer()
Form6.Height = 8172
Form6.Width = 14868
If Image2.Width < 10000 Then Image2.Width = Image2.Width + Int((Rnd() + 0.5) * (Rnd() + 0.5) * (Rnd() + 1) * 100)
If Image2.Width > 10000 Then
    ShowCursor True
    a = MsgBox("对不起，您安装的flash版本较低，无法进入教程关卡。您可以选择放弃教程，直接进入老版（2.4.7）主菜单（不过画质会降低很多）。或者重试。", (5 + 48), "警告")
    If a = 4 Then
    Else
        nmal = True
        Unload Form6
        Form7.Visible = True
    End If
End If
End Sub

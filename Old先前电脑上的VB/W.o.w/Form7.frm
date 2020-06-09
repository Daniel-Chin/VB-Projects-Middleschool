VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "主菜单"
   ClientHeight    =   6192
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8568
   LinkTopic       =   "Form7"
   ScaleHeight     =   6192
   ScaleWidth      =   8568
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer3 
      Left            =   3360
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   720
   End
   Begin VB.Image Image10 
      Height          =   480
      Left            =   4560
      Picture         =   "Form7.frx":0000
      Top             =   5520
      Width           =   3240
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   600
      Picture         =   "Form7.frx":0FB3
      Top             =   5520
      Width           =   3240
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "公测3.0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   372
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Image Image8 
      Height          =   900
      Left            =   5760
      Picture         =   "Form7.frx":1CCC
      Top             =   0
      Width           =   2796
   End
   Begin VB.Image Image7 
      Height          =   1824
      Left            =   2760
      Picture         =   "Form7.frx":3C1B
      Top             =   2160
      Visible         =   0   'False
      Width           =   3336
   End
   Begin VB.Image Image6 
      Height          =   1824
      Left            =   2760
      Picture         =   "Form7.frx":89E5
      Top             =   2160
      Visible         =   0   'False
      Width           =   3336
   End
   Begin VB.Image Image5 
      Height          =   864
      Left            =   6960
      Picture         =   "Form7.frx":D9AF
      Top             =   3720
      Visible         =   0   'False
      Width           =   864
   End
   Begin VB.Image Image4 
      Height          =   864
      Left            =   3840
      Picture         =   "Form7.frx":ECC1
      Top             =   3720
      Visible         =   0   'False
      Width           =   864
   End
   Begin VB.Image Image3 
      Height          =   864
      Left            =   720
      Picture         =   "Form7.frx":FFB5
      Top             =   3720
      Visible         =   0   'False
      Width           =   864
   End
   Begin VB.Image Image2 
      Height          =   2112
      Left            =   0
      Picture         =   "Form7.frx":1124C
      Top             =   3000
      Width           =   12228
   End
   Begin VB.Image Image1 
      Height          =   6216
      Left            =   0
      Picture         =   "Form7.frx":1C48A
      Top             =   0
      Width           =   8604
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Dim time2, time3 As Long
Dim nmal As Boolean
Private Sub Form_Unload(Cancel As Integer)
If nmal = True Then
Else
    End
End If
End Sub
Private Sub Form_Load()
ShowCursor True
Form11.WindowsMediaPlayer1.URL = App.Path & "\.W.o.w_files\dark.mp3"
nmal = False
Image2.Height = 15
Image3.Height = 15
Image4.Height = 15
Image5.Height = 15
Image9.Width = 15
Image10.Width = 15
End Sub
Private Sub Image3_Click()   'maoxian
Image6.Visible = True
End Sub
Private Sub Image4_Click()
Image7.Visible = True
End Sub

Private Sub Image5_Click()
Timer1.Interval = 0
nmal = True
Unload Form7
Form8.Visible = True
End Sub

Private Sub Image6_Click()
Image6.Visible = False
End Sub

Private Sub Image7_Click()
Image7.Visible = False
End Sub

Private Sub Timer1_Timer()
Form7.Height = 6732
Form7.Width = 8784
End Sub

Private Sub Timer2_Timer()
time2 = time2 + 1
If time2 > 50 Then Timer2.Interval = 0: Timer3.Interval = 50: Image3.Visible = True: Image4.Visible = True: Image5.Visible = True
Image2.Height = Image2.Height + Int(2112 / 50 + 1)
Image9.Width = Image9.Width + Int(3240 / 50 + 1)
Image10.Width = Image10.Width + Int(3240 / 50 + 1)
End Sub

Private Sub Timer3_Timer()
time3 = time3 + 1
If time3 > 20 Then Timer3.Interval = 0
Image3.Height = Image3.Height + 60
Image4.Height = Image4.Height + 60
Image5.Height = Image5.Height + 60
End Sub

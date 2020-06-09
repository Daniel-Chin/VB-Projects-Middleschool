VERSION 5.00
Begin VB.Form Launcher 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "我去冒险"
   ClientHeight    =   3816
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8592
   Icon            =   "Launcher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   3816
   ScaleWidth      =   8592
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Waiter 
      Interval        =   3000
      Left            =   3720
      Top             =   2520
   End
   Begin VB.Image Im 
      Height          =   372
      Left            =   1320
      Picture         =   "Launcher.frx":0CCA
      Stretch         =   -1  'True
      Top             =   840
      Width           =   972
   End
End
Attribute VB_Name = "Launcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Size As Single = 0.9, SF As Single = 0.6

Private Sub Form_Load()
    Me.Width = Screen.Width * SF
    Me.Height = Me.Width / 859 * 381
    Im.Move Width * (1 - Size) / 2, Height * (1 - Size) / 2, Width * Size, Height * Size
    Show
    SaveSetting "iGlope", "woqu", "run", "1"
End Sub

Private Sub Waiter_Timer()
Shell App.Path & "\bin\files\wqmx.exe", vbNormalFocus
Unload Me
End Sub

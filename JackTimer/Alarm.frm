VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Alarm 
   AutoRedraw      =   -1  'True
   Caption         =   "JackTimer"
   ClientHeight    =   5916
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10572
   Icon            =   "Alarm.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5916
   ScaleWidth      =   10572
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8400
      Top             =   3000
   End
   Begin VB.Timer Main 
      Interval        =   10000
      Left            =   7080
      Top             =   3600
   End
   Begin WMPLibCtl.WindowsMediaPlayer P 
      Height          =   2892
      Left            =   3480
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   2892
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5101
      _cy             =   5101
   End
End
Attribute VB_Name = "Alarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Form_Load()
Tag = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '-1 Do it,-2 Reset
Me.BackColor = 0
Main.Interval = 10000
Me.FontSize = 166
Me.ForeColor = vbWhite
Print "时间到！"
With P
    .settings.volume = 100
    .URL = App.Path & "\1.mp3"
    .Controls.play
End With
End Sub

Private Sub Main_Timer()
Beep
Me.BackColor = vbRed - Me.BackColor
Cls
Print "时间到！"
P.Controls.currentPosition = 0
P.Controls.play
End Sub

Private Sub Timer1_Timer()
Me.Caption = P.Controls.currentPosition
End Sub

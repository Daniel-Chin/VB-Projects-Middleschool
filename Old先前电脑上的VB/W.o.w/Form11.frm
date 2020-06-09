VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   2340
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   3624
   LinkTopic       =   "Form11"
   ScaleHeight     =   2340
   ScaleWidth      =   3624
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   2880
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1320
      Top             =   960
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   2412
      Left            =   360
      TabIndex        =   0
      Top             =   0
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
      _cy             =   4255
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim grace As Integer
Public a_b As Boolean
Dim a As Integer
Private Sub Form_Load()
WindowsMediaPlayer1.settings.volume = 100
WindowsMediaPlayer1.Left = -WindowsMediaPlayer1.Width - 1
WindowsMediaPlayer1.Top = -WindowsMediaPlayer1.Height - 1
WindowsMediaPlayer1.URL = App.Path & "\.W.o.w_files\music.mp3"
a_b = False
End Sub

Private Sub Timer1_Timer()
If a_b = False Then
    If WindowsMediaPlayer1.playState = wmppsStopped And a = 0 Then
        WindowsMediaPlayer1.URL = App.Path & "\.W.o.w_files\dark.mp3"
        WindowsMediaPlayer1.Controls.play
        a = 1
    End If
    If WindowsMediaPlayer1.playState = wmppsStopped And a = 1 Then
        WindowsMediaPlayer1.URL = App.Path & "\.W.o.w_files\ahahah.mp3"
        WindowsMediaPlayer1.Controls.play
    End If
Else
    If WindowsMediaPlayer1.Controls.currentPosition > 7.2 Then
        WindowsMediaPlayer1.Controls.currentPosition = 1
    End If
End If
End Sub

Private Sub Timer2_Timer()
Form11.Visible = False
Timer2.Interval = 0
End Sub

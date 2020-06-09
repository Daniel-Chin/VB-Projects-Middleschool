VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form10 
   BackColor       =   &H80000008&
   Caption         =   "新闻联播"
   ClientHeight    =   6468
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   10404
   LinkTopic       =   "Form10"
   ScaleHeight     =   6468
   ScaleWidth      =   10404
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4440
      Top             =   1320
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   6492
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10452
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
      _cx             =   18436
      _cy             =   11451
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nmal As Boolean
Dim time As Long
Dim a As Integer
Private Sub Form_Load()
Form11.WindowsMediaPlayer1.Controls.pause
WindowsMediaPlayer1.URL = App.Path & "\.W.o.w_files\新闻联播.wmv"
WindowsMediaPlayer1.Controls.play
nmal = False
End Sub

Private Sub Timer1_Timer()
If Form10.WindowsMediaPlayer1.Controls.currentPosition > 83.5 Then Form10.WindowsMediaPlayer1.Controls.pause: a = MsgBox("节选结束了。再看一遍？", vbOKCancel, "再次播放？")
If a = 1 Then WindowsMediaPlayer1.Controls.currentPosition = 0: WindowsMediaPlayer1.Controls.play: time = 0: a = 0
If a = 2 Then nmal = True: Unload Form10: Form5.Visible = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
If nmal = True Then
    Form11.WindowsMediaPlayer1.Controls.play
Else
    End
End If
End Sub

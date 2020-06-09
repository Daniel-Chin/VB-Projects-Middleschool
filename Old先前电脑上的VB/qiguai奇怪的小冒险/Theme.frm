VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Theme 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.Timer Tm 
      Interval        =   100
      Left            =   2640
      Top             =   600
   End
   Begin WMPLibCtl.WindowsMediaPlayer P 
      Height          =   612
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   1812
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   5
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
      _cx             =   3196
      _cy             =   1080
   End
End
Attribute VB_Name = "Theme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TongGuan As Boolean

Private Sub Form_Load()
With P
    .URL = App.Path & "\R\1.wma"
    .Settings.volume = 100
End With
End Sub

Private Sub Tm_Timer()
If P.Controls.currentPosition >= 45 Then
    P.Controls.currentPosition = 0
End If
If Settings.MusicOff Then
    P.Controls.stop
Else
    P.Controls.play
End If
End Sub

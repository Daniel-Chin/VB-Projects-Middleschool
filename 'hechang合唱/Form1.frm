VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   7050
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer1 
      Left            =   3120
      Top             =   3480
   End
   Begin WMPLibCtl.WindowsMediaPlayer P 
      Height          =   2410
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   2530
      URL             =   "C:\Users\wx\Documents\main2.wma"
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
      _cx             =   4463
      _cy             =   4251
   End
   Begin WMPLibCtl.WindowsMediaPlayer P 
      Height          =   2410
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   2530
      URL             =   "C:\Users\wx\Documents\main1.2.wma"
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
      _cx             =   4463
      _cy             =   4251
   End
   Begin WMPLibCtl.WindowsMediaPlayer P 
      Height          =   2410
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   -480
      Width           =   2530
      URL             =   "C:\Users\wx\Documents\a1.wma"
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
      _cx             =   4463
      _cy             =   4251
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Boolean

Private Sub Form_Load()
P(0).URL = "C:\Users\wx\Documents\b1.wma"
P(1).URL = "C:\Users\wx\Documents\b2.wma"
P(2).URL = "C:\Users\wx\Documents\b3.wma"
Load P(3)
P(3).URL = "C:\Users\wx\Documents\b4.wma"
Load P(4)
P(4).URL = "C:\Users\wx\Documents\b5.wma"
Load P(5)
P(5).URL = "C:\Users\wx\Documents\b6.wma"
'Load P(6)
'P(6).URL = "C:\Users\wx\Documents\a0.wma"
Timer1.Interval = 1000
End Sub

Private Sub Timer1_Timer()
If Not a Then
    a = True
    For k = 0 To 5
        P(k).Controls.pause
        P(k).Controls.currentPosition = 0
    Next k
Else
    For k = 0 To 5
        P(k).Controls.play
    Next k
Timer1 = False
End If
End Sub

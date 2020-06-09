VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form haha 
   Caption         =   "Form1"
   ClientHeight    =   1840
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   2980
   LinkTopic       =   "Form1"
   ScaleHeight     =   1840
   ScaleWidth      =   2980
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   720
   End
   Begin WMPLibCtl.WindowsMediaPlayer P 
      Height          =   1810
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   2530
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
      _cx             =   4463
      _cy             =   3193
   End
End
Attribute VB_Name = "haha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Interval = 24
P.URL = "C:\Users\wx\Documents\d.wma"
End Sub

Private Sub Timer1_Timer()
If P.Controls.currentPosition >= 26 Then
    P.Controls.currentPosition = 0.5
End If
End Sub

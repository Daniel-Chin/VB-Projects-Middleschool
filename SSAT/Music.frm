VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Music 
   Caption         =   "SSAT Music Player"
   ClientHeight    =   4716
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   Icon            =   "Music.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4716
   ScaleWidth      =   3624
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.Timer T 
      Interval        =   6
      Left            =   2400
      Top             =   2640
   End
   Begin WMPLibCtl.WindowsMediaPlayer P 
      Height          =   2172
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   3132
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
      _cx             =   5525
      _cy             =   3831
   End
   Begin WMPLibCtl.WindowsMediaPlayer P 
      Height          =   2172
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3132
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
      _cx             =   5525
      _cy             =   3831
   End
End
Attribute VB_Name = "Music"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Mode As String
Dim GB As Byte
Dim Toler As Integer
Const Ttp As Integer = 666

Public Sub ModeSwitch(ModeName As String)
Mode = ModeName
Select Case Mode
    Case "menu"
        P(0).URL = App.Path & "\0.wma"
    Case "boss"
        P(1).Controls.stop
        P(0).URL = App.Path & "\btls.wma"
        P(0).Controls.currentPosition = 36.1
    Case "lv"
        P(1).Controls.stop
        P(0).URL = App.Path & "\bdf.wma"
    Case "fps"
        P(1).Controls.stop
        P(0).URL = App.Path & "\fps.wma"
End Select
End Sub

Private Sub Form_Load()
Randomize
P(0).settings.volume = 100
P(1).settings.volume = 100
T.Interval = 8000
End Sub

Private Sub T_Timer()
Dim Ass As Single, NewUrl As String
Select Case Mode
    Case "menu"
        Ass = 10.7
        NewUrl = App.Path & "\" & Int(Rnd() * 4) & ".wma"
    Case "boss"
        Ass = 135
        NewUrl = App.Path & "\btls.wma"
    Case "lv"
        Ass = 73.6
        NewUrl = App.Path & "\bdf.wma"
    Case "fps"
        Ass = 10.8
        NewUrl = App.Path & "\fps.wma"
End Select
If P(GB).Controls.currentPosition >= Ass Then
GT:
    Toler = 0
    GB = 1 - GB
    P(GB).URL = NewUrl
    P(GB).Controls.play
Else
    If P(GB).Controls.currentPosition = 0 Then
        T.Interval = 66
        Toler = Toler + T.Interval
        If Toler >= Ttp + T.Interval Then
            GoTo GT
        End If
    Else
        Toler = 0
        T.Interval = Abs((Ass - P(GB).Controls.currentPosition) * 466) + 1
    End If
End If
End Sub

VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5184
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5184
   ScaleWidth      =   9780
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Exiter 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6960
      Top             =   2760
   End
   Begin VB.Timer Starter 
      Interval        =   1000
      Left            =   7800
      Top             =   2160
   End
   Begin WMPLibCtl.WindowsMediaPlayer P 
      Height          =   2892
      Index           =   0
      Left            =   3480
      TabIndex        =   0
      Top             =   1200
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim Tms As Integer
Dim FileName(0 To 9999) As String
Dim FileNum As Integer

Dim K As Integer

Dim Root As String

Private Sub Exiter_Timer()
K = 0
Do While P(K).Controls.currentPosition = 0
    K = K + 1
    If K > FileNum Then
        Unload Me
        End
    End If
Loop
End Sub

Private Sub Form_Load()
Root = App.Path & "\"
Me.Visible = False
Beep
Dim B As String
B = "0.mp3"
With P(0)
    .URL = Root & B
    .settings.volume = 0
    .Visible = True
    .Controls.play
End With
End Sub

Private Sub SK(KEYS As String)
Dim myKey As Object
Set myKey = CreateObject("WScript.Shell")
myKey.SendKeys KEYS
End Sub

Private Sub Starter_Timer()
If Tms = 3 Then
    Starter = False
    SK "%s"
    For K = 0 To FileNum
        With P(K)
            .settings.volume = 100
            .Controls.currentPosition = 0
            .Controls.play
        End With
    Next K
    Exiter = True
Else
    Beep
    Tms = Tms + 1
End If
End Sub

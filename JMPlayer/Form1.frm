VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JMPlayer"
   ClientHeight    =   1008
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4752
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1008
   ScaleWidth      =   4752
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Chker 
      Interval        =   66
      Left            =   2520
      Top             =   480
   End
   Begin WMPLibCtl.WindowsMediaPlayer P 
      Height          =   972
      Left            =   2040
      TabIndex        =   1
      Top             =   120
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
      _cy             =   1715
   End
   Begin VB.Label RP 
      BackColor       =   &H00000000&
      Caption         =   "Replay"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1008
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2306
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Dim Started As Boolean

Private Sub Chker_Timer()
If P.Controls.currentPosition = 0 Then
    If Started Then
        Unload Me
        End
    End If
Else
    Started = True
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 119 'w
        P.settings.volume = P.settings.volume + 3
    Case 115 's
        P.settings.volume = P.settings.volume - 3
    Case Else
        Exit Sub
End Select
Caption = String(P.settings.volume / 10, "#")
End Sub

Private Sub Form_Load()
Dim CD As String, tT As Integer
CD = Command
tT = InStr(CD, "once")
If tT Then
    Chker = True
    CD = Left(CD, tT - 1)
Else
    Chker = False
End If
a = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
Width = RP.Width + 66
Height = RP.Height + 406
tT = InStr(CD, Chr(34))
If tT Then
    CD = Right(CD, Len(CD) - tT)
    tT = InStr(CD, Chr(34))
    If tT Then
        CD = Left(CD, tT - 1)
    End If
End If
With P
    .URL = CD
    .settings.volume = 100
    .Controls.play
End With
End Sub

Private Sub RP_Click()
P.Controls.currentPosition = 0
P.Controls.play
End Sub

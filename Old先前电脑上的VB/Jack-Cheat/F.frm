VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form F 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11388
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   11388
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Starter 
      Interval        =   10
      Left            =   8160
      Top             =   1920
   End
   Begin VB.Timer fBianSe 
      Left            =   5640
      Top             =   1920
   End
   Begin VB.Timer Main 
      Interval        =   20
      Left            =   4560
      Top             =   2520
   End
   Begin VB.Timer Miscer2 
      Left            =   5880
      Top             =   2520
   End
   Begin VB.Timer Miscer 
      Left            =   5400
      Top             =   2640
   End
   Begin VB.Timer Clicker 
      Left            =   3960
      Top             =   2520
   End
   Begin VB.Timer BianSe 
      Left            =   1320
      Top             =   960
   End
   Begin WMPLibCtl.WindowsMediaPlayer Misc 
      Height          =   732
      Left            =   600
      TabIndex        =   0
      Top             =   -1000
      Visible         =   0   'False
      Width           =   1452
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
      _cx             =   2561
      _cy             =   1291
   End
   Begin VB.Line RB 
      BorderWidth     =   4
      X1              =   3120
      X2              =   3360
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line LB 
      BorderWidth     =   4
      X1              =   3120
      X2              =   2880
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line RT 
      BorderWidth     =   4
      X1              =   3360
      X2              =   3120
      Y1              =   2520
      Y2              =   2280
   End
   Begin VB.Line LT 
      BorderWidth     =   4
      X1              =   2880
      X2              =   3120
      Y1              =   2520
      Y2              =   2280
   End
   Begin VB.Line Z 
      BorderWidth     =   2
      X1              =   3120
      X2              =   3120
      Y1              =   2280
      Y2              =   2760
   End
   Begin VB.Line H 
      BorderWidth     =   2
      X1              =   2880
      X2              =   3360
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label L 
      Caption         =   "杰克17代-正在配置"
      BeginProperty Font 
         Name            =   "华文隶书"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   2028
      TabIndex        =   1
      Top             =   3084
      Width           =   7332
   End
End
Attribute VB_Name = "F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim St As Single, sTd As Byte
Dim BC As Byte, BCt As Byte
Dim Xx As Integer, Yy As Integer
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Dim Can As Boolean
Const BsXs As Byte = 60
Dim MainR As Byte, MainG As Byte, MainB As Byte, YsCase As Byte
Private Sub BianSe_Timer()
Select Case YsCase
    Case 0
        MainR = MainR + BsXs
        If MainR = 240 Then YsCase = 1 'R
    Case 1
        MainG = MainG + BsXs
        If MainG = 240 Then YsCase = 2 'RG
    Case 2
        MainR = MainR - BsXs
        If MainR = 0 Then YsCase = 3 'G
    Case 3
        MainB = MainB + BsXs
        If MainB = 240 Then YsCase = 4 'GB
    Case 4
        MainG = MainG - BsXs
        If MainG = 0 Then YsCase = 5 'B
    Case 5
        MainR = MainR + BsXs
        If MainR = 240 Then YsCase = 6 'BR
    Case 6
        MainB = MainB - BsXs
        If MainB = 0 Then YsCase = 1 'R
End Select
H.BorderColor = RGB(MainR, MainG, MainB)
Z.BorderColor = RGB(MainR, MainG, MainB)
LT.BorderColor = RGB(MainR, MainG, MainB)
LB.BorderColor = RGB(MainR, MainG, MainB)
RT.BorderColor = RGB(MainR, MainG, MainB)
RB.BorderColor = RGB(MainR, MainG, MainB)
'
L.ForeColor = RGB(255 - MainR, 255 - MainG, 255 - MainB)
End Sub


Private Sub fBianSe_Timer()
BCt = BCt + 1
BC = Sin(BCt / 5) * 30 + 65
Me.BackColor = RGB(0, 0, BC)
L.BackColor = RGB(0, 0, BC)
If BCt >= 30 Then
    BCt = 0
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 27     'ESC
        ShowCursor True
        End
End Select
End Sub

Private Sub Form_Load()
Randomize
ShowCursor False
St = 4.5
BC = 45 '定义背景色
Misc.URL = App.Path & "\Main2.mp3"
If Rnd() > 0.5 Then
    Misc.URL = App.Path & "\Main.mp3"
    Miscer.Interval = 20
Else
    Miscer2.Interval = 20
End If
Misc.Controls.play
Me.BackColor = RGB(0, 0, BC)
L.ForeColor = RGB(0, 0, BC)
L.BackColor = RGB(0, 0, BC)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xx = X
Yy = Y
End Sub

Private Sub L_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(0, 0, X + L.Left, Y + L.Top)
End Sub

Private Sub Main_Timer()
Dim Shang As Integer, Xia As Integer, Zuo As Integer, You As Integer
Shang = Yy - 240
Xia = Yy + 240
Zuo = Xx - 240
You = Xx + 240
With Z
    .X1 = Xx
    .X2 = Xx
    .Y1 = Shang
    .Y2 = Xia
End With
With H
    .X1 = Zuo
    .X2 = You
    .Y1 = Yy
    .Y2 = Yy
End With
With LT
    .X1 = Xx
    .X2 = Zuo
    .Y1 = Shang
    .Y2 = Yy
End With
With LB
    .X1 = Zuo
    .X2 = Xx
    .Y1 = Yy
    .Y2 = Xia
End With
With RT
    .X1 = Xx
    .X2 = You
    .Y1 = Shang
    .Y2 = Yy
End With
With RB
    .X1 = You
    .X2 = Xx
    .Y1 = Yy
    .Y2 = Xia
End With
End Sub

Private Sub Miscer2_Timer()
If Misc.Controls.currentPosition > 35.1 Then
    Misc.Controls.currentPosition = 18.2
End If
End Sub

Private Sub Miscer_Timer()
If Misc.Controls.currentPosition > 18.1 Then
    Misc.Controls.currentPosition = 10
End If
End Sub

Private Sub Click()
Clicker.Interval = 20
Can = False
End Sub

Private Sub Starter_Timer()
St = St + 0.05
sTd = sTd + 1
F.Height = (Sin(St) + 1) * 3960
F.Top = Screen.Height / 2 - F.Height / 2
F.Width = Screen.Width - (Sin(St) + 1) * 2000
F.Left = Screen.Width / 2 - F.Width / 2
F.BackColor = RGB(255 - 255 / 68 * sTd, 255 - 255 / 68 * sTd, 100 - 55 / 68 * sTd)
L.BackColor = F.BackColor
L.ForeColor = F.BackColor
If St > 7.9 Then
    Starter.Interval = 0
    fBianSe.Interval = 50
    BianSe.Interval = 50
End If
End Sub

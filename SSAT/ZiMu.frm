VERSION 5.00
Begin VB.Form ZiMu 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "我去冒险"
   ClientHeight    =   4944
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10596
   Icon            =   "ZiMu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4944
   ScaleWidth      =   10596
   Begin VB.Timer Exiter 
      Enabled         =   0   'False
      Interval        =   1666
      Left            =   6840
      Top             =   3000
   End
   Begin VB.Timer Main 
      Left            =   4800
      Top             =   2280
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   5352
      TabIndex        =   0
      Top             =   1440
      Width           =   108
   End
   Begin VB.Line S 
      Index           =   0
      Visible         =   0   'False
      X1              =   4800
      X2              =   5760
      Y1              =   2280
      Y2              =   2640
   End
End
Attribute VB_Name = "ZiMu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Dim k As Long, KK As Integer
Dim Sy As Long
Const FxLays As Integer = 6
Dim FxNum(0 To 6) As Integer
'Dim PixPerCm As Double

Private Sub Exiter_Timer()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    Sy = 16
Else
    Sy = -16
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Sy = -6
End Sub

Private Sub Form_Load()
Tag = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '-1 Do it,-2 Reset
Sy = -6
Randomize
Move 0, 0, Screen.Width, Screen.Height
L.Move 0, Height + 666, Width, Height '
L.FontSize = 36
FxNum(0) = 1
Dim N As Integer
N = 66 '
For k = 1 To FxLays
    FxNum(k) = FxNum(k - 1) + N
    N = Int(N * 0.36) '
Next k
Main.Interval = 6 '
Dim KK As Integer
Dim R As Byte
For KK = 1 To FxLays
    For k = FxNum(KK - 1) To FxNum(KK) - 1
        Load S(k)
        With S(k)
            .BorderColor = RGB(255 - Int(Rnd() ^ 2 * 172), 255 - Int(Rnd() ^ 2 * 19), 255)
            If Rnd <= 0.006 Then
                .BorderColor = vbBlue
            End If
            If Rnd <= 0.006 Then
                .BorderColor = RGB(255, 255, 0)
            End If
            .BorderWidth = KK + 1
            .X1 = Rnd * Width
            .X2 = .X1
            .Y1 = Rnd * Height
            .Y2 = .Y1
            .Visible = True
        End With
    Next k
Next KK
Main_Timer
LPrt "我去冒险"
LPrt ""
LPrt ""
LPrt "玩家您好！"
LPrt "我是这个游戏的作者。"
LPrt "我叫iGlope。"
LPrt "我来自高一(5)班。"
LPrt ""
LPrt ""
LPrt "内测人员"
LPrt ""
LPrt "Azure"
LPrt "大鼻子"
LPrt "Jessica"
LPrt "KIM"
LPrt "莫笛坤"
LPrt "o龙"
LPrt "竹子"
LPrt ""
LPrt ""
LPrt "特别鸣谢"
LPrt ""
LPrt "本游戏的背景音乐"
LPrt "由我、贝多芬和披头士乐队创作"
LPrt ""
LPrt "游戏灵感来自"
LPrt "Only This Level 3"
LPrt "Don't shoot the puppy"
LPrt "The Stanley's Parable"
LPrt "猫里奥"
LPrt "I Wanna be the Guy"
LPrt ""
LPrt "感谢万能的 Visual Basic 6.0"
LPrt ""
LPrt "Thank you for playing!"
LPrt "现在是" & Now
If Month(Now) = 10 And Day(Now) = 1 Then
    LPrt "中华人民共和国万岁！"
End If
LPrt ""
LPrt ""
LPrt "我去冒险"
'Show
'Me.ScaleMode = 7
'PixPerCm = Me.Width / Me.ScaleWidth
'Me.ScaleMode = 1
End Sub

Private Sub Main_Timer()
KK = KK Mod FxLays
KK = KK + 1
For k = FxNum(KK - 1) To FxNum(KK) - 1
    With S(k)
        .Y1 = .Y1 + Sy * KK
        If .Y1 > Height + 100 Then .Y1 = -100: .X1 = Rnd * Width: .X2 = .X1
        If .Y1 < -100 Then .Y1 = Height + 100: .X1 = Rnd * Width: .X2 = .X1
        .Y2 = .Y1 - Sy * KK * 1.6
    End With
Next k
If KK = 1 Then
    If L.Height + L.Top >= 1966 Then
        L.Top = L.Top + Sy * (FxLays)
    Else
        Exiter = True
    End If
End If
End Sub

Private Sub LPrt(Text As String)
L = L & Text & Chr(10)
End Sub

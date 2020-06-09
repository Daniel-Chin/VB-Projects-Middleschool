VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Gloper"
   ClientHeight    =   6504
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11052
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6504
   ScaleWidth      =   11052
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox FirstP 
      BackColor       =   &H00400000&
      Caption         =   "第一次玩"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   1932
   End
   Begin VB.Timer Loadder 
      Interval        =   24
      Left            =   10560
      Top             =   2160
   End
   Begin VB.Timer E3 
      Interval        =   83
      Left            =   6600
      Top             =   720
   End
   Begin VB.Timer E2 
      Left            =   6120
      Top             =   840
   End
   Begin VB.Timer E1 
      Left            =   5400
      Top             =   720
   End
   Begin VB.Timer Quieter 
      Left            =   9240
      Top             =   2280
   End
   Begin VB.Timer DJ 
      Interval        =   10
      Left            =   8760
      Top             =   3480
   End
   Begin VB.Timer Exiter 
      Left            =   9720
      Top             =   3480
   End
   Begin VB.Timer Ens 
      Left            =   8280
      Top             =   2520
   End
   Begin VB.Timer HuanChongQi 
      Left            =   7560
      Top             =   3480
   End
   Begin VB.Timer BianSe 
      Left            =   6720
      Top             =   3840
   End
   Begin VB.Label Starter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "极限模式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   70.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1572
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   6972
   End
   Begin VB.Label Starter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "经典模式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   70.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1572
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   6972
   End
   Begin WMPLibCtl.WindowsMediaPlayer Music 
      Height          =   972
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   -2000
      Visible         =   0   'False
      Width           =   852
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
      _cx             =   1503
      _cy             =   1715
   End
   Begin VB.Shape En 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0FF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   612
      Index           =   3
      Left            =   6480
      Top             =   4800
      Width           =   4332
   End
   Begin VB.Shape En 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0FF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   1812
      Index           =   2
      Left            =   2160
      Top             =   4200
      Width           =   972
   End
   Begin VB.Shape En 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0FF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   1212
      Index           =   1
      Left            =   7920
      Top             =   240
      Width           =   1932
   End
   Begin VB.Shape En 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0FF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   1092
      Index           =   0
      Left            =   2760
      Top             =   840
      Width           =   1932
   End
   Begin VB.Shape Player 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      FillStyle       =   0  'Solid
      Height          =   1092
      Left            =   4800
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Board 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   100.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4920
      TabIndex        =   0
      Top             =   4080
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim k As Byte

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Type BlockEnemy
    Top_Speed As Integer
    Left_Speed As Integer
    Color As Byte
    Speed As Integer
End Type
Dim Life As Integer, wL As Integer
Const TopLife As Byte = 200
Dim E(0 To 3) As BlockEnemy

Dim Tms As Long, Lj As Byte
Dim LoadTms As Byte

Dim GameOn As Boolean, Holding As Boolean

Dim HuanChong As Boolean

Dim YsCase As Byte, MainR As Byte, MainB As Byte, MainG As Byte
Const BsXs As Byte = 20, BackGround As Byte = 30
Dim Zz As Boolean

Dim Ex As Integer

Private Sub Board_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Holding Then
    ShowCursor False
    Holding = True
    If Not GameOn Then  '开局
        Reset
        GameOn = True
        BianSe.Interval = 50
        Ens.Interval = 19
    End If
End If
Board_MouseMove Button, Shift, X, Y
End Sub

Private Sub Board_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If GameOn = True And Holding = True And HuanChong = False Then
    With Player
        .Top = Y / Me.Height * (Me.Height - .Height)
        .Left = X / Me.Width * (Me.Width - .Width)
        If .Top = 0 Or Me.Height - .Top - .Height <= 30 Then
            Die
        End If
        If .Left = 0 Or Me.Width - .Left - .Width <= 30 Then
            Die
        End If
    End With
    HuanChong = True
    HuanChongQi.Interval = 18
End If
End Sub

Private Sub Board_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Holding And GameOn Then
    ShowCursor True
    Holding = False
End If
End Sub

Private Sub DJ_Timer()
If Music(-Zz).Controls.currentPosition >= 6.6 Then
    With Music(1 + Zz)
        .URL = App.Path & "\music.wma"
        .Controls.currentPosition = 0
        .Controls.play
    End With
    Zz = Not Zz
    DJ.Interval = 1000
Else
    DJ.Interval = Abs((6.6 - Music(-Zz).Controls.currentPosition) * 466) + 1
    Debug.Print DJ.Interval
End If
End Sub

Private Sub E3_Timer()
If Month(Now) = 5 And Day(Now) = 24 Then
Else
    E3.Interval = 0
    Exit Sub
End If
If E1.Interval = 0 Then
    E1.Interval = 1000
Else
    E2.Interval = 1000
    E3.Interval = 0
End If
End Sub

Private Sub E1_Timer()
Me.Caption = "Happy Birthday! "
End Sub

Private Sub E2_Timer()
Me.Caption = "Gloper"
End Sub

Private Sub Ens_Timer() '敌军行动
If GameOn Then
    Tms = Tms + 1
    Lj = (Lj + 1) Mod 10
    For k = 0 To 3
        With E(k)
            If Lj = 0 Or Tms <= 80 Then
                .Speed = .Speed + 2
            End If
            En(k).Top = En(k).Top + .Top_Speed * .Speed
            En(k).Left = En(k).Left + .Left_Speed * .Speed
            If En(k).Top <= 0 Or En(k).Top + En(k).Height >= Me.Height Then
                .Top_Speed = -.Top_Speed
            End If
            If En(k).Left <= 0 Or En(k).Left + En(k).Width >= Me.Width Then
                .Left_Speed = -.Left_Speed
            End If
            If JieChu(En(k).Top, En(k).Height, En(k).Left, En(k).Width) Then
                '碰到
                .Speed = Int(.Speed * 0.9)
                Life = Life - wL * 20
                If Life <= 0 Then
                    Life = 0
                    Die
                    Exit Sub
                End If
                .Color = .Color + 5
                If .Color >= 55 Then
                    .Color = 55
                End If
                En(k).FillColor = RGB(255 - .Color * 3, 200 - .Color * 3, 200)
                En(k).BorderColor = En(k).FillColor
            Else
                If .Color <> 0 Then
                    .Color = .Color - 1
                    En(k).FillColor = RGB(255 - .Color * 3, 200 - .Color, 200)
                    En(k).BorderColor = En(k).FillColor
                End If
            End If
            If .Speed >= 300 Then
                '高速加分
                Tms = Tms * 1.002
                En(k).FillColor = vbRed
            End If
        End With
    Next k
End If
End Sub

Private Sub Exiter_Timer()
Ex = Ex + 1
Me.Top = Ex ^ 2 * 5
If Screen.Height - 2 * Me.Top <= 0 Then
    Me.Visible = False
    Exiter.Interval = 0
    Exit Sub
End If
Me.BackColor = Me.BackColor + RGB(5, 5, 0)
Me.Height = Screen.Height - 2 * Me.Top
SetLayeredWindowAttributes hwnd, 0, 200 - Ex * 6, 2
End Sub

Private Sub FirstP_KeyUp(KeyCode As Integer, Shift As Integer)
Form_KeyUp KeyCode, Shift
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27     'ESC
        If Holding Then
            ShowCursor True
        End If
        TuiChu
End Select
End Sub

Private Sub Form_Load()
Randomize
SetWindowLongA hwnd, -20, 524288
SetLayeredWindowAttributes hwnd, 0, 1, 2
With Board
    .Top = 0
    .Left = 0
    .Height = Screen.Height
    .Width = Screen.Width
    .Visible = False
    .FontSize = 90
    .Caption = Chr(10) & "Gloper"
End With
With Music(0)
    .Controls.play
    .settings.volume = 100
    .URL = App.Path & "\music.wma"
End With
For k = 0 To 3
    With En(k)
        .Visible = False
    End With
Next k
Player.Visible = False
Load Music(1)
Music(1).Top = -2000
Reset
With Starter(0)
    .Left = (Screen.Width - .Width) / 2
    .Top = Screen.Height / 5
End With
With Starter(1)
    .Left = (Screen.Width - .Width) / 2
    .Top = Screen.Height / 2
End With
FirstP.Value = GetSetting("Gloper") '''''''''''
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For k = 0 To 1
    With Starter(k)
        .ForeColor = vbBlack
        .BackColor = vbWhite
    End With
Next k
End Sub

Private Sub HuanChongQi_Timer()
HuanChong = False
HuanChongQi.Interval = 0
End Sub

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
Player.FillColor = RGB(MainR, MainG, MainB)
Player.BorderColor = RGB(255 - MainR, 255 - MainG, 255 - MainB)
'Healer
Life = Life + 3
If Life > TopLife Then
    Life = TopLife
End If
GxCl
'打分
If GameOn Then
    Board = Tms
End If
'加分
Select Case Tms
    Case 10 ^ 3 To 10 ^ 4
        Tms = Tms + 4
    Case 10 ^ 4 To 10 ^ 5
        Tms = Tms + 40
    Case 10 ^ 5 To 10 ^ 6
        Tms = Tms + 200
    Case 10 ^ 6 To 10 ^ 7
        Tms = Tms + 1000
    Case Is > 10 ^ 7
        Tms = Tms + 8000
End Select
End Sub

Private Sub Die()
Life = 0
GxCl
Ens.Interval = 0
GameOn = False
BianSe.Interval = 0
If Holding Then
    ShowCursor True
    Holding = False
End If
Dim a As String, b As Integer
b = Len(Str(Tms)) - 2
Select Case b
    Case 0
        a = "朽木"
    Case 1
        a = "粪墙"
    Case 2
        a = "可雕玉"
    Case 3
        a = "紫水晶"
    Case 4
        a = "黄金。"
    Case 5
        a = "钻石。"
    Case 6
        a = "小石像！"
    Case Is >= 7
        a = "神探詹姆斯！！！"
End Select
Board = Chr(10) & "评分：" & Left(Str(Tms), IIf(b > 7, b - 5, 2)) & "阶" & a
MsgBox "你死了。", vbOKOnly + vbCritical, "Gloper"
Reset
End Sub

Private Sub Reset()
Life = TopLife
Tms = 0
GameOn = False
Me.BackColor = RGB(0, 0, BackGround)
Dim LL(), Tt()
LL = Array(1, -1, 1, -1)
Tt = Array(1, 1, -1, -1)
For k = 0 To 3
    E(k).Left_Speed = LL(k)
    E(k).Top_Speed = Tt(k)
    E(k).Color = 0
    E(k).Speed = 1
    En(k).FillColor = RGB(255, 200, 200)
    En(k).BorderColor = En(k).FillColor
Next k
Dim jG As Integer, zHb As Byte  '纵横比
zHb = 4
With Screen
    jG = .Height / 12
    En(0).Top = jG
    En(0).Left = jG * zHb
    En(1).Top = jG
    En(1).Left = .Width - jG * zHb - En(1).Width
    En(2).Top = .Height - jG - En(2).Height
    En(2).Left = jG * zHb
    En(3).Top = .Height - jG - En(3).Height
    En(3).Left = .Width - jG * zHb - En(3).Width
End With
With Player
    .Top = (Screen.Height - .Height) / 2
    .Left = (Screen.Width - .Width) / 2
End With
With Me
    .Top = 0
    .Left = 0
    .Width = Screen.Width
    .Height = Screen.Height
    .ForeColor = vbWhite
End With
With Board
    .BackColor = RGB(0, 0, BackGround)
    .ForeColor = vbWhite
End With
End Sub

Private Sub GxCl()  '更新背景颜色
Me.BackColor = RGB(TopLife - Life, 0, BackGround)
Board.BackColor = Me.BackColor
End Sub

Private Function JieChu(Top As Long, Height As Long, Left As Long, Width As Long) As Boolean
With Player
    If .Top <= Top + Height And .Top + .Height >= Top Then
        If .Left <= Left + Width And .Left + .Width >= Left Then
            JieChu = True
        End If
    End If
End With
End Function

Private Sub TuiChu()
'退出
Reset
Board.Visible = False
For k = 0 To 3
    En(k).Visible = False
Next k
Starter(0).Visible = False
Starter(1).Visible = False
Player.Visible = False
BianSe.Interval = 0
Ens.Interval = 0
Exiter.Interval = 30
Quieter.Interval = 80
End Sub

Private Sub Loadder_Timer()
SetLayeredWindowAttributes hwnd, 0, LoadTms, 2
If LoadTms >= 166 Then
    Loadder.Interval = 0
    Exit Sub
End If
LoadTms = LoadTms + 15
End Sub

Private Sub Quieter_Timer()
Music(0).settings.volume = Music(0).settings.volume - 7
Music(1).settings.volume = Music(1).settings.volume - 7
Quieter.Interval = Quieter.Interval + 8
If Music(1).settings.volume <= 20 Then
    Quieter.Interval = Quieter.Interval + 100
End If
If Music(1).settings.volume <= 1 Then
    End
End If
End Sub

Private Sub Starter_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
With Starter(index)
    .BackColor = vbBlack
    .ForeColor = vbWhite
End With
With Starter(1 - index)
    .BackColor = vbWhite
    .ForeColor = vbBlack
End With
End Sub

Private Sub Starter_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If index = 0 Then
    wL = 1
Else
    If Button = 1 Then
        wL = 3
    Else
        MsgBox "你进入了无敌模式。", vbOKOnly + vbInformation, "Gloper"
        wL = 0
    End If
End If
Board.Visible = True
For k = 0 To 3
    With En(k)
        .Visible = True
    End With
Next k
Player.Visible = True
Starter(0).Visible = False
Starter(1).Visible = False
FirstP.Visible = False
End Sub

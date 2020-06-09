VERSION 5.00
Begin VB.Form LvDog 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "偷西瓜"
   ClientHeight    =   7164
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10452
   DrawWidth       =   4
   FillColor       =   &H00008000&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7164
   ScaleWidth      =   10452
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Twi 
      Enabled         =   0   'False
      Interval        =   336
      Left            =   7680
      Top             =   4200
   End
   Begin VB.Timer Main 
      Interval        =   36
      Left            =   8160
      Top             =   3120
   End
   Begin VB.Timer Dogger 
      Interval        =   333
      Left            =   6960
      Tag             =   "0"
      Top             =   3360
   End
   Begin VB.PictureBox Dog 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   3240
      ScaleHeight     =   492
      ScaleWidth      =   732
      TabIndex        =   0
      Top             =   2280
      Width           =   732
   End
   Begin VB.Line Rope 
      BorderWidth     =   6
      Visible         =   0   'False
      X1              =   0
      X2              =   3120
      Y1              =   360
      Y2              =   0
   End
   Begin VB.Shape Player 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   1560
      Top             =   360
      Width           =   372
   End
   Begin VB.Shape Trunk 
      FillStyle       =   0  'Solid
      Height          =   348
      Left            =   5076
      Shape           =   3  'Circle
      Top             =   3600
      Width           =   312
   End
End
Attribute VB_Name = "LvDog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Pi As Double = 3.14159265358979
Dim oX As Long, oY As Long
Dim r As Long
Dim CuAn As Double, RoL As Long
Dim vX As Long, vY As Long
Dim pW As Boolean, pA As Boolean, pS As Boolean, pD As Boolean
Const F As Long = 6
Const mR As Long = 200
Dim MLeft As Byte

Private Type Point
    X As Long
    Y As Long
End Type

Private Type MPoint
    X() As Variant
    Y() As Variant
End Type

Dim Mellon As MPoint
Dim Lamda As Double

Private Sub Main_Timer()
If vX Or vY Or pA Or pW Or pS Or pD Then
    vX = vX + (pA - pD) * F
    vY = vY + (pW - pS) * F
    vX = vX * 0.98
    If vX >= 0 Then
        vX = Abss(vX - F / 6)
    Else
        vX = Abss(vX + F / 6, False)
    End If
    vY = vY * 0.98
    If vY >= 0 Then
        vY = Abss(vY - F / 6)
    Else
        vY = Abss(vY + F / 6, False)
    End If
    With Player
        If .Left <= 0 Then
            .Left = 1
            vX = -vX / 2
        End If
        If .Left >= Width - .Width Then
            .Left = Width - .Width - 1
            vX = -vX / 2
        End If
        If .Top <= 0 Then
            .Top = 1
            vY = -vY / 2
        End If
        If .Top >= Height - .Height Then
            .Top = Height - .Height - 1
            vY = -vY / 2
        End If
        If Sqr((Center(Player).X - oX) ^ 2 + (Center(Player).Y - oY) ^ 2) <= r * 4 Then
            vX = -vX * 2
            vY = -vY * 2
        End If
        .Left = .Left + vX
        .Top = .Top + vY
    End With
    Dim k As Byte
    For k = 0 To 15
        If Sqr((Center(Player).X - Mellon.X(k)) ^ 2 + (Center(Player).Y - Mellon.Y(k)) ^ 2) <= mR * 2 Then
            Circle (Mellon.X(k), Mellon.Y(k)), mR + 1, vbWhite
            Mellon.X(k) = -666
            Mellon.Y(k) = -666
            MLeft = MLeft - 1
            If MLeft = 0 Then
                Win
            Else
                Twi.Tag = 8
                Twi.Enabled = True
                Twi_Timer
            End If
        End If
    Next k
    Tag = MODx(ShAn - CuAn)
    Dim Delta As Double
    If Tag <= Pi Then
        Delta = Tag
    Else
        Delta = Tag - 2 * Pi
    End If
    If CuAn + Delta > 31.1765181709071 Then
        Delta = 31.177 - CuAn
    End If
    CuAn = CuAn + Delta
    RoL = RoL - Delta * r * Min(CuAn ^ 2 / 366, 6)
    With Rope
        .X2 = r * Cos(CuAn) + oX
        .Y2 = oY - r * Sin(CuAn)
        .X1 = .X2 + RoL * Cos(CuAn + Pi / 2)
        .Y1 = .Y2 - RoL * Sin(CuAn + Pi / 2)
    End With
    With Dog
        .Left = Rope.X1 - .Width / 2
        .Top = Rope.Y1 - .Height / 2
    End With
End If
Dim PX As Long, PY As Long
Dim DX As Long, DY As Long
PX = Center(Player).X
PY = Center(Player).Y
DX = Center(Dog).X
DY = Center(Dog).Y
Tag = Sqr((PX - DX) ^ 2 + (PY - DY) ^ 2)
If Tag <= 606 Then
    Si
End If
Dogger.Interval = Tag / 12
End Sub

Private Sub Dog_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub Dog_KeyUp(KeyCode As Integer, Shift As Integer)
Form_KeyUp KeyCode, Shift
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 38 'W
        pW = True
    Case 37 'A
        pA = True
    Case 40 'S
        pS = True
    Case 39 'D
        pD = True
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 38 'W
        pW = False
    Case 37 'A
        pA = False
    Case 40 'S
        pS = False
    Case 39 'D
        pD = False
End Select '
End Sub

Private Sub Dogger_Timer()
Dogger.Tag = 1 - Dogger.Tag
With Dog
    .Cls
    .PaintPicture LoadPicture(App.Path & "\dog" & Dogger.Tag & ".jpg"), 0, 0, .Width, .Height
End With
End Sub

Private Sub Form_Load()
Dogger_Timer
Reset
End Sub

Private Function Center(Thing As Object) As Point
With Thing
    Center.X = .Left + .Width / 2
    Center.Y = .Top + .Height / 2
End With
End Function

Private Function Abss(Number, Optional Mode As Boolean = True)
If Mode Then
    If Number <= 0 Then
        Abss = 0
    Else
        Abss = Number
    End If
Else
    If Number >= 0 Then
        Abss = 0
    Else
        Abss = Number
    End If
End If
End Function

Private Function ShAn() As Double
Dim A As Double
Dim PlayerX As Long, PlayerY As Long
PlayerX = Center(Player).X - oX
PlayerY = oY - Center(Player).Y
If PlayerX = 0 Then PlayerX = 1
If PlayerX = r Then PlayerX = r + 1
If PlayerX = -r Then PlayerX = 1 - r
Tag = r * Sqr(PlayerX ^ 2 + PlayerY ^ 2 - r ^ 2)
A = Atn((PlayerX * PlayerY + Tag) / (PlayerX ^ 2 - r ^ 2))
PlayerX = Center(Player).X
PlayerY = Center(Player).Y
ShAn = MODx(A - Pi / 2)
If PlayerY <= oY Then
    If PlayerX <= oX + r Then
        ShAn = ShAn + Pi
    End If
Else
    If PlayerX <= oX - r Then
        ShAn = ShAn + Pi
    End If
End If
End Function

Function MODx(A As Double, Optional B As Double = 2 * Pi) As Double
If B > 0 Then
    Do Until A >= 0 And A < B
        If A < 0 Then
            A = A + B
        Else
            A = A - B
        End If
    Loop
    MODx = A
Else
    Error 17
End If
End Function

Function Min(A, B)
If A >= B Then
    Min = B
Else
    Min = A
End If
End Function

Private Sub Si()
If MsgBox("你被那条被西瓜田主人拴在西瓜田中央的一根木桩上用来看守西瓜田可后来却被那个很戳的程序员画成这副戳样的凶猛黑色纯种德国牧羊犬咬死了。", vbCritical + vbRetryCancel, "偷西瓜") = vbRetry Then
    Dim You As Long
    You = 1
    Do While MsgBox("你" & String(You, "又") & "重新试着去X戏那条狗，于是你又死了。", vbCritical + vbRetryCancel, "重试的结果") = vbRetry
        You = You + 1
        If You = 336 Then
            MsgBox "溢出！", vbCritical + vbOKOnly, "Error"
            Exit Do
        End If
    Loop
End If
Reset
End Sub

Private Sub Win()
MsgBox "你通过了奖励关卡“偷西瓜”。", vbOKOnly + vbInformation, "恭喜！"
Unload Me
End
End Sub

Private Sub Reset()
Me.Width = 10524
With Player
    .Left = 9486
    .Top = 966
    .FillStyle = 1
End With
With Dog
    .Left = 3246
    .Top = 2286
End With
With Rope
    .Visible = False
    .X1 = Center(Dog).X
    .Y1 = Center(Dog).Y
    .X2 = Center(Trunk).X
    .Y2 = Center(Trunk).Y
    .Visible = True
End With
vX = -1: vY = 0
pW = False
pA = False
pS = False
pD = False
Cls
Me.FillColor = &H8000&
Dim k As Byte
With Mellon
    .X = Array(7526, 6182, 3221, 4049, 8688, 483, 8202, 4484, 9296, 6129, 5599, 570, 5000, 6642, 2814, 8852)
    .Y = Array(4039, 2192, 5886, 3660, 5368, 3135, 5985, 4283, 425, 4256, 5808, 4486, 2257, 4985, 2115, 6243)
    For k = 0 To 15
        Me.Circle (.X(k), .Y(k)), mR, vbGreen
    Next k
End With
Me.FillColor = vbWhite
CuAn = Pi
RoL = 2866
oX = Center(Trunk).X
oY = Center(Trunk).Y
r = Trunk.Width / 3
Twi.Enabled = False
MLeft = 16
Me.Caption = "偷西瓜"
Main_Timer
End Sub

Private Sub Twi_Timer()
If Twi.Tag Mod 2 = 1 Then
    Tag = Format(MLeft, "00")
Else
    Tag = "   "
End If
If Twi.Tag = 0 Then
    Twi = False
Else
    Twi.Tag = Twi.Tag - 1
    Me.Caption = "偷西瓜 - 还剩" & Tag & "个西瓜"
End If
End Sub

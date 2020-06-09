VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "奶奶游戏"
   ClientHeight    =   5652
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9252
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5652
   ScaleWidth      =   9252
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Hider 
      Interval        =   1000
      Left            =   4200
      Top             =   2640
   End
   Begin VB.Timer Bf 
      Left            =   4560
      Top             =   3960
   End
   Begin VB.Timer Main 
      Left            =   5400
      Top             =   3000
   End
   Begin VB.Label Block 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   2040
      Width           =   612
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tms As Byte
Dim Sc As Boolean

Private Declare Function ShowCursor Lib "user32" (ByVal BShow As Long) As Long

Dim Dx As Integer, dXs As Integer
Dim Speed As Integer
Const Pi As Double = 3.14155265358575
Dim Dir(1 To 35) As Integer

Private Sub Bf_Timer()
For i = 1 To 35
    Dir(i) = (Dir(i) + Int((Rnd() - 0.5) * 10) + 360) Mod 360
Next i
Speed = Speed + Int(Rnd() * 5) - 2
dXs = dXs + Int(Rnd() * 5) - 2
Dx = Dx + dXs
If dXs < -15 Then dXs = -15
If dXs > 20 Then dXs = 20
If Abs(Speed) > 1000 Then
    Reset
End If
If Dx <= 0 Or Dx > 2000 Then
    Reset
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 27     'ESC
        ShowCursor True
        End
    Case Else
        Reset
End Select
End Sub

Private Sub Form_Load()
ShowCursor False
Sc = False
Randomize
With Me
    .Top = 0
    .Left = 0
    .Height = Screen.Height
    .Width = Screen.Width
End With
With Block(0)
    .Height = Me.Height / 5
    .Width = Me.Width / 7
    .Top = -.Height
    .Left = -.Width
End With
For i = 1 To 35
    Load Block(i)
Next i
Reset
End Sub

Private Function RndCl() As Long
Select Case Int(Rnd() * 4) + 1
    Case 1
        RndCl = vbRed
    Case 2
        RndCl = vbGreen
    Case 3
        RndCl = vbBlue
    Case 4
        RndCl = vbWhite
End Select
End Function

Private Sub Reset()
Dx = 500
dXs = 10
For i = 1 To 5 * 7
    With Block(i)
        .Top = ((i - 1) Mod 5) * Block(0).Height + Dx
        .Left = ((i - 1) \ 5) * Block(0).Width + Dx
        .Height = Dx
        .Width = Dx
        .BackColor = RndCl
        .Visible = True
    End With
    Dir(i) = 0
Next i
Main.Interval = 50
Bf.Interval = 100
Speed = 100
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Rnd() < 0.2 Then
    Reset
    Tms = 0
    If Not Sc Then
        Sc = True
        ShowCursor True
    End If
End If
End Sub

Private Sub Hider_Timer()
Tms = Tms + 1
If Tms > 5 Then
    Tms = 0
    If Sc Then
        Sc = False
        ShowCursor flase
    End If
End If
End Sub

Private Sub Main_Timer()
For i = 1 To 7 * 5
    With Block(i)
        .Top = (.Top + Sin(Dir(i) * Pi / 180) * Speed + Me.Height) Mod Me.Height
        .Left = (.Left + Cos(Dir(i) * Pi / 180) * Speed + Me.Width) Mod Me.Width
        .Height = Dx
        .Width = Dx
    End With
Next i
End Sub

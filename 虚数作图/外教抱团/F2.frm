VERSION 5.00
Begin VB.Form F2 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer MainT 
      Interval        =   66
      Left            =   1320
      Top             =   960
   End
End
Attribute VB_Name = "F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Dim TMD As Byte
Dim dTMD As Integer

Const TpTMD As Byte = 200
Const Size As Single = 0.4
Const TpSpd As Long = 66
Const Brt As Byte = 50
Const SpdTMD As Integer = 3

Dim SX As Long, SY As Long, SXX As Long, SYY As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
Main.Form_KeyPress KeyAscii
End Sub

Private Sub Form_Load()
Randomize
TMD = 0
MainT.Interval = 66
Dim a
a = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
Show
Me.Visible = False
SetWindowLongA hwnd, -20, 524288
SetLayeredWindowAttributes hwnd, 0, 200, 2
dTMD = SpdTMD
End Sub

Private Sub MainT_Timer()
If TMD = 0 Then
    'Reset
    Me.Visible = False
    With Screen
        Move Rnd() * .Width, Rnd() * .Height, Rnd() * .Width * Size, Rnd() * .Height * Size
    End With
    Me.BackColor = RGB(Int(Rnd() * (256 - Brt) + Brt), Int(Rnd() * (256 - Brt) + Brt), Int(Rnd() * (256 - Brt) + Brt))
    SX = Rnd() * TpSpd - TpSpd / 2
    SY = Rnd() * TpSpd - TpSpd / 2
    SXX = Rnd() * TpSpd * 2 - TpSpd
    SYY = Rnd() * TpSpd * 2 - TpSpd
    dTMD = SpdTMD
    TMD = dTMD
    SetLayeredWindowAttributes hwnd, 0, 0, 2
    Me.Visible = True
Else
    If TMD >= TpTMD Then
        dTMD = -dTMD
    End If
    TMD = dTMD + TMD
    With Me
        If .Left + .Width <= 0 Or .Left >= Screen.Width Or .Top + .Height <= 0 Or .Top >= Screen.Height Then
            TMD = 0
            Exit Sub
        End If
        If .Height + SYY <= 0 Then
            SY = SY + SYY
            SYY = -SYY
        End If
        If .Width + SXX <= 0 Then
            SX = SX + SXX
            SXX = -SXX
        End If
        .Move .Left + SX, .Top + SY, .Width + SXX, .Height + SYY
    End With
    SetLayeredWindowAttributes hwnd, 0, TMD, 2
End If
End Sub

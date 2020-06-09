VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Main 
      Interval        =   36
      Left            =   2160
      Top             =   1320
   End
   Begin VB.Line s 
      Index           =   0
      X1              =   3360
      X2              =   3120
      Y1              =   1920
      Y2              =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor Lib "user32" (ByVal BShow As Long) As Long

Dim SX As Long, SY As Long
Dim pX As Integer, pY As Integer
Const AMN As Integer = 66

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ShowCursor True: End
End Sub

Private Sub Form_Load()
ShowCursor False
Main.Interval = 36
With Me
    .Move 0, 0, Screen.Width, Screen.Height
    .BackColor = 0
End With
For K = 1 To AMN
    Load s(K)
    With s(K)
        .BorderColor = RGB(256 - Rnd ^ 2 * 256, 255, 255)
        .BorderWidth = 1 + Rnd ^ 6 * 6
        If K Mod 2 = 0 Then
            .BorderWidth = 1
        Else
            If .BorderWidth = 1 Then .BorderWidth = 2
        End If
        .X1 = Rnd * Width
        .X2 = .X1
        .Y1 = Rnd * Height
        .Y2 = .Y1
        .Visible = True
    End With
Next K
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pX <> 0 Then
    SX = SX + (X - pX) / 5
    SY = SY + (Y - pY) / 5
End If
pX = X
pY = Y
End Sub

Private Sub Main_Timer()
SX = SX * 0.95
SY = SY * 0.95
Dim bS As Double
Dim sT As Integer
sT = 1
Hey:
For K = sT To AMN Step 2
    With s(K)
        bS = Sqr((.X1 - Width / 2) ^ 2 + (.Y1 - Height / 2) ^ 2) / Sqr(Width ^ 2 + Height ^ 2) * 2
        If bS >= 1 Then bS = 0.99999
        bS = Cos(Atn(bS / Sqr(-bS * bS + 1)))
        .X1 = .X1 + SX * bS
        If .X1 > Width + 100 Then .X1 = -100: .Y1 = Rnd * Height
        If .X1 < -100 Then .X1 = Width + 100: .Y1 = Rnd * Height
        .Y1 = .Y1 + SY * bS
        If .Y1 > Height + 100 Then .Y1 = -100: .X1 = Rnd * Width
        If .Y1 < -100 Then .Y1 = Height + 100: .X1 = Rnd * Width
        .X2 = .X1 + SX * bS
        .Y2 = .Y1 + SY * bS
        'Debug.Print bS
    End With
Next K
SX = -SX
SY = -SY
If sT = 1 Then
    sT = 2
    GoTo Hey
End If
End Sub

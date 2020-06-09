VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7944
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7944
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   2160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowCursor Lib "user32" (ByVal BShow As Long) As Long
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Const DX As Long = 524
Dim W As Long, H As Long
Dim DN As Boolean
Dim Tms As Integer

Dim S(0 To 99, 0 To 99) As Byte

Private Sub Form_Load()
ShowCursor False
Tag = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '-1 Do it,-2 Reset
Move 0, 0, Screen.Width, Screen.Height
With Me
    .FillStyle = 0
    .BackColor = 0
    W = .Width \ DX + 1
    H = .Height \ DX + 1
End With
Timer1.Interval = 6
Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 4 Then DN = True: Tms = 0
If Button = 2 Then DN = True: Tms = -524
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DN = False
End Sub

Private Sub Timer1_Timer()
If DN Then
    If Tms = 66 Then
        ShowCursor True
        Unload Me
        End
    End If
    Tms = Tms + 1
Else
    Tms = 0
End If

Dim a As Long, b As Long, c As Long
a = Int(Rnd() * W)
b = Int(Rnd() * H)
S(a, b) = Int(Rnd() * 256)
Dim aa As Long, bb As Long
Dim aa1 As Long, aa2 As Long, bb1 As Long, bb2 As Long
aa1 = a - 1
aa2 = a + 1
Select Case a
    Case 0
        aa1 = 0
    Case W
        aa2 = W
End Select
bb1 = b - 1
bb2 = b + 1
Select Case b
    Case 0
        bb1 = 0
    Case H
        bb2 = H
End Select
For aa = aa1 To aa2
    For bb = bb1 To bb2
        c = PJ(aa, bb)
        Me.ForeColor = RGB(0, 0, c)
        Me.FillColor = RGB(0, 0, c)
        Line (aa * DX, bb * DX)-(aa * DX + DX, bb * DX + DX), , B
    Next bb
Next aa
End Sub

Private Function PJ(a As Long, b As Long) As Long
Dim aa As Long, bb As Long
Dim aa1 As Long, aa2 As Long, bb1 As Long, bb2 As Long
aa1 = a - 1
aa2 = a + 1
Select Case a
    Case 0
        aa1 = 0
    Case W
        aa2 = W
End Select
bb1 = b - 1
bb2 = b + 1
Select Case b
    Case 0
        bb1 = 0
    Case H
        bb2 = H
End Select
Dim JS As Integer
For aa = aa1 To aa2
    For bb = bb1 To bb2
        JS = JS + 1
        PJ = PJ + S(aa, bb)
    Next bb
Next aa
PJ = Int(PJ / JS)
PJ = FD(PJ)
End Function

Private Function FD(a As Long) As Long
Dim b As Double
b = a
FD = (Sin((b - 128) / 256 * 3.14159265358979) + 1) * 128
End Function

Sub fuck()
ShowCursor True
End Sub

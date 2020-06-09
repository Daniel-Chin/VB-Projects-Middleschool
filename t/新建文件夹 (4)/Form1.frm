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
Dim a As Long, b As Long, c As Long
For a = 0 To W
    For b = 0 To H
        c = RGB(0, 0, Int(Rnd() * 256))
        Me.ForeColor = c
        Me.FillColor = c
        Line (a * DX, b * DX)-(a * DX + DX, b * DX + DX), , B
    Next b
Next a
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
a = Int(Rnd() * W) * DX
b = Int(Rnd() * H) * DX
c = RGB(0, 0, Int(Rnd() * 256))
Me.ForeColor = c
Me.FillColor = c
Line (a, b)-(a + DX, b + DX), , B
End Sub

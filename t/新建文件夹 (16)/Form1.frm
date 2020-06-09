VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5556
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5904
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5556
   ScaleWidth      =   5904
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Pi As Double = 3.14159265358979

Private Sub Form_Load()
Show
Dim Hei, Wid
Move 0, 0, 6000, 6000
Hei = Me.ScaleHeight
Wid = Me.ScaleWidth
Dim k
Me.DrawWidth = 2
For k = 0 To 2 * Pi Step Pi / 100000
    Me.ForeColor = CL(k)
    Line (Wid / 2, Hei / 2)-(Cos(k) * Wid / 2 + Wid / 2, Sin(k) * Wid / 2 + Hei / 2)
Next k
SavePicture Me.Image, "d:\LG.bmp"
Unload Me
End Sub

Function CL(k) As Long
Dim t
t = k
Dim r, g, b
r = 0
g = 0
b = 0
Select Case t
    Case 0 To Pi / 3 * 2
        r = 1 - 3 / Pi / 2 * t
        g = 3 / Pi / 2 * t
    Case Pi / 3 * 2 To Pi / 3 * 4
        g = 2 - 3 / Pi / 2 * t
        b = -1 + 3 / Pi / 2 * t
    Case Else
        b = 3 - 3 / Pi / 2 * t
        r = -2 + 3 / Pi / 2 * t
End Select
CL = RGB(Int(r * 256), Int(g * 256), Int(b * 256))
End Function

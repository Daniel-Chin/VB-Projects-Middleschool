VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6324
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   11556
   LinkTopic       =   "Form1"
   ScaleHeight     =   6324
   ScaleWidth      =   11556
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Complex
    Re As Single
    Im As Single
End Type

Const Pi = 3.14159265358979

Dim A(0 To 32) As Complex
Dim B(0 To 32) As Complex

Private Type Camera
    XForm As Long
    YForm As Long
    Stretch As Single
End Type

Dim Cam As Camera

Private Function cTimes(Z1 As Complex, Z2 As Complex) As Complex
cTimes.Re = Z1.Re * Z2.Re - Z1.Im * Z2.Im
cTimes.Im = Z1.Re * Z2.Im + Z1.Im * Z2.Re
End Function

Private Function P2C(Re As Single, Im As Single) As Complex
P2C.Re = Re
P2C.Im = Im
End Function

Private Function C2txt(Z As Complex) As String
C2txt = Z.Re
Select Case Z.Im
    Case Is > 0
        C2txt = C2txt & "+" & Z.Im & "i"
    Case Is < 0
        C2txt = C2txt & Z.Im & "i"
End Select
End Function

Private Sub Graph(Data() As Complex, Colour As Long)
Dim k As Integer
Dim PreX As Long, PreY As Long, X As Long, Y As Long
PreX = Cam.XForm + Data(0).Re * Cam.Stretch
PreY = Cam.YForm - Data(0).Im * Cam.Stretch
ForeColor = Colour
For k = 1 To UBound(Data)
    X = Cam.XForm + Data(k).Re * Cam.Stretch
    Y = Cam.YForm - Data(k).Im * Cam.Stretch
    Line (X, Y)-(PreX, PreY)
    PreX = X
    PreY = Y
Next k
Line (0, Cam.YForm)-(Me.ScaleWidth, Cam.YForm)
Line (Cam.XForm, 0)-(Cam.XForm, Me.ScaleHeight)
Line (Cam.XForm + Cam.Stretch, Cam.YForm - Cam.Stretch / 5)-(Cam.XForm + Cam.Stretch, Cam.YForm + Cam.Stretch / 5)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim CX As Single, CY As Single
CX = (X - Cam.XForm) / Cam.Stretch
CY = -(Y - Cam.YForm) / Cam.Stretch
Dim k As Integer
For k = 0 To 32
    A(k).Re = 2 * Cos(2 * Pi * k / 32) + CX
    A(k).Im = 2 * Sin(2 * Pi * k / 32) + CY
    B(k) = cTimes(A(k), A(k))
Next k
Cls
Graph A, vbWhite
Graph B, RGB(255, 255, 0)
End Sub

Private Sub SetCam(X As Long, Y As Long, Stretch As Single)
With Cam
    .XForm = X
    .YForm = Y
    .Stretch = Stretch
End With
End Sub

Private Sub Form_Resize()
SetCam Me.ScaleWidth / 2, Me.ScaleHeight / 2, 1000
End Sub

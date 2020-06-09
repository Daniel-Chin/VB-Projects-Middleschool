VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00012345&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1464
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3156
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1464
   ScaleWidth      =   3156
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    Main.Form_KeyPress KeyAscii
End Sub

Sub Form_Load()
    SetWindowLong hwnd, -20, GetWindowLong(hwnd, -20) Or &H80000
    SetLayeredWindowAttributes hwnd, &H12345, 0, &H1
End Sub

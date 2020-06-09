VERSION 5.00
Begin VB.Form AnNiu 
   BackColor       =   &H00012345&
   BorderStyle     =   0  'None
   ClientHeight    =   1464
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3156
   LinkTopic       =   "AnNiu0"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1464
   ScaleWidth      =   3156
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "AnNiu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    Main.Form_KeyPress KeyAscii
End Sub

Sub Form_Load()
    SetWindowLong hwnd, -20, GetWindowLong(hwnd, -20) Or &H80000
    SetLayeredWindowAttributes hwnd, &H12345, 0, &H1
    Me.Tag = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)  '-1为置前，-2
End Sub

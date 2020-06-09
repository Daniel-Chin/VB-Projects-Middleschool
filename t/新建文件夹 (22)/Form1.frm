VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   3744
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Const Speed As Integer = 50

Private Sub K(Keys As String)
Dim myKey As Object
Set myKey = CreateObject("WScript.Shell")
myKey.SendKeys Keys
Sleep Speed
DoEvents
'Ctrl    ^
'Shift   +
'Alt     %
End Sub

Private Sub Form_Load()
Open "D:\vbbreak" For Binary As #1
Put #1, 1, 0
Close #1
Beep
Sleep 3000
Beep
Sleep 3000
Beep
Dim C As Byte
Do
    DoEvents
    Open "D:\vbbreak" For Binary As #1
    Get #1, 1, C
    If C <> 0 Then
        Put #1, 1, 0
        Close #1
        End
    End If
    Close #1
    K "{home}"
    K "+{end}"
    Dian 1
    Dian 4
    Dian 2
    Dian 3
    K "{DOWN}{DOWN}"
Loop
End Sub

Sub Dian(whatTo As Integer)
Sleep Speed
Select Case whatTo
    Case 1
        SetCursorPos 470, 140
        mouse_event 6, 0, 0, 0, 0
    Case 4
        SetCursorPos 470, 300
        mouse_event 6, 0, 0, 0, 0
    Case 2
        SetCursorPos 660, 187
        mouse_event 6, 0, 0, 0, 0
    Case 3
        SetCursorPos 660, 265
        mouse_event 6, 0, 0, 0, 0
    Case Else
        Stop
End Select
End Sub

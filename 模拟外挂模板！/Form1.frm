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

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Sub Snd(ByVal f As Long, Optional Time As Long = 100)
Beep f, Time
DoEvents
End Sub

Private Sub Music(Note As Integer)
Snd 440 * 2 ^ ((Note + 2) / 12)
End Sub

Private Sub Form_Load()
Open "D:\vbbreak" For Binary As #1
End Sub

Private Sub K(Txt As String)
Dim myKey As Object
Set myKey = CreateObject("WScript.Shell")
myKey.SendKeys Txt
'Ctrl    ^
'Shift   +
'Alt %
End Sub

Private Sub Dian()
mouse_event 6, 0, 0, 0, 0
End Sub

Private Sub Zhi(X As Long, Y As Long)
SetCursorPos X, Y
End Sub

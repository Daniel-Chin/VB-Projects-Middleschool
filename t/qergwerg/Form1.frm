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

Const Speed As Integer = 200

Private Sub Snd(ByVal f As Long, Optional Time As Long = 100)
Beep f, Time
DoEvents
End Sub

Private Sub Music(Note As Integer)
Snd 440 * 2 ^ ((Note + 2) / 12)
End Sub

Private Sub Form_Load()
Open "D:\vbbreak" For Binary As #1
Dim C As Byte
Put #1, 1, 0
Music 1
Sleep 1000
Music 5
Sleep 1000
Music 8
Sleep 1000
Music 13
K "^a"
K "^c"
K "^{tab}"
K "^k"
Sleep 1000
Dian 400, 320
K "xjz{tab}"
K "xjz{tab}"
K "159159qaz{tab}"
K "159159qaz{tab}"
K "xjzlnxh59@163.com{tab}"
K "xjzlnxh59@163.com{tab}"
K "^v{tab}"
K "{tab}"
K "^+{tab}"
K "{tab}"
K "^c^{tab}"
K Mid(Clipboard.GetText, 7, 4) & "{tab}"
K Val(Mid(Clipboard.GetText, 11, 2)) & "{tab}"
K Val(Mid(Clipboard.GetText, 13, 2)) & "{tab}"
K "{down}{tab}"
K "{tab}"
K "{enter}"
Dian 813, 362
Dian 812, 300
Dian 279, 443
Dian 516, 702
K "{tab}{tab}"
K "^v{tab}"
K "{down}{down}{down}{down}{down}{down}{down}{down}{down}{down}{down}{down}{down}{down}{down}{down}{tab}"
K "{down}{tab}"
K "{tab}{tab}{tab}{tab}{tab}"
K "{enter}"
Dian 808, 297
Dian 260, 388
Dian 281, 321
Dian 505, 703
K "{tab}{tab}"
K "{tab}{tab}"
K "{down}{down}{down}{down}{down}{down}{down}{down}{down}{tab}"
K "{down}{down}{down}{down}{down}{down}{down}{down}{down}{down}{down}{down}{down}{tab}"
K "{tab}{tab}{tab}{tab}{tab}"
K " "
K "+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}"
DoEvents
Close #1
End
End Sub

Private Sub K(Txt As String)
Dim C As Byte
Get #1, 1, C
If C = 1 Then
    Close #1
    End
End If
Dim myKey As Object
Set myKey = CreateObject("WScript.Shell")
myKey.SendKeys Txt
'Ctrl    ^
'Shift   +
'Alt %
Sleep Speed
End Sub

Private Sub Dian(X As Long, Y As Long)
SetCursorPos X, Y
Sleep 20
mouse_event 6, 0, 0, 0, 0
Sleep Speed * 2
End Sub

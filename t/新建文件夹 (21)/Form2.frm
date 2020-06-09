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
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Sub Snd(ByVal f As Long, Optional Time As Long = 100)
Beep f, Time
DoEvents
End Sub

Private Sub Form_Load()
Show
Open "d:\vbbreak" For Binary As #1
Put #1, 1, 0
Snd 440
Sleep 1000
Snd 440
Sleep 1000
Snd 440
Sleep 1000
Snd 440
For j = 1 To 26
    Snd 440: K "{ENTER}"
    Sleep 300
    Snd 880: K "^e"
    Snd 440: K "1194"
    Snd 880: K "{ENTER}"
    Snd 440: K "^s"
    K "%{F4}"
    Sleep 200
    K "{RIGHT}"
Next j
End Sub

Sub K(txtK As String)
    DoEvents
    Dim C As Byte
    Get #1, 1, C
    If C = 1 Then
        Unload Me
        End
    End If
Dim myKey As Object
Set myKey = CreateObject("WScript.Shell")
myKey.SendKeys txtK
DoEvents
Wait
End Sub

Sub Wait()
Sleep 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #1
End Sub


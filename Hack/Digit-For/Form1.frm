VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   1320
      Top             =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Dim nuMm As Integer

Private Sub Snd(ByVal f As Long, Optional Time As Long = 100)
Beep f, Time
DoEvents
End Sub
Private Sub Form_Load()
Me.Visible = False
Timer1.Interval = 36
Timer1 = False
Snd 1000
Sleep 1000
Snd 1100
Sleep 1000
Snd 1300
Sleep 1000
Snd 1500
Snd 1500, 500
Timer1 = True
End Sub

Private Sub K(A As String)
Dim myKey As Object
Set myKey = CreateObject("WScript.Shell")
myKey.SendKeys A
End Sub

Private Sub Timer1_Timer()
nuMm = nuMm + 1
K "{Enter}"
K "+{Tab}"
K Format(nuMm, "0000")
K "{Enter}"
If nuMm = 9999 Then
    Snd 2000
    Snd 2000, 1000
    End
End If
End Sub


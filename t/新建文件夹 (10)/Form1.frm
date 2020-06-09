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
      Interval        =   1000
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
Dim a As Integer
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Sub Snd(ByVal f As Long, Optional Time As Long = 100)
Beep f, Time
DoEvents
End Sub

Private Sub Form_Load()
Me.Visible = False
Timer1.Interval = 360
Timer1 = False
Snd 1000
Sleep 1000
Snd 900
Sleep 1000
Snd 800
Sleep 1000
Snd 700
Snd 700
Timer1 = True
End Sub

Private Sub Timer1_Timer()
Dim myKey As Object
Set myKey = CreateObject("WScript.Shell")
myKey.SendKeys "{Esc}"
myKey.SendKeys Format(a, "00")
myKey.SendKeys "{Enter}"
a = a + 1
If a = 100 Then End
End Sub

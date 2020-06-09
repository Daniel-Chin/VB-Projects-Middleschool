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

Const W As Integer = 9
Const L As Integer = 6
Dim M(1 To L) As Integer

Private Sub Snd(ByVal f As Long, Optional Time As Long = 100)
Beep f, Time
DoEvents
End Sub
Private Sub Form_Load()
Reset
Me.Visible = False
Timer1.Interval = 1
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
Timer1.Interval = 3
Do
    gb = 1
    If M(1) = W Then
        Do Until M(gb) <> W
            M(gb) = 0
            gb = gb + 1
            If gb > L Then Snd 900: Snd 900, 1000: End
        Loop
        M(gb) = M(gb) + 1
    Else
        M(1) = M(1) + 1
    End If
    For kk = 1 To L
'        K Chr(M(kk) + 65)
'        Debug.Print Chr(M(kk) + 65);
'''''''''''''''''''''''
        K Chr(M(kk) + 48)
        Debug.Print Chr(M(kk) + 48);
    Next kk
    Debug.Print ""
'    K "{Enter}"
    K "{Esc}"
Open "D:\vbbreak" For Binary As #1
Dim C As Byte
Get #1, 1, C
If C = 1 Then Stop
Close #1
    DoEvents
Loop
Tag = Val(Tag) + 1
If Val(Tag) = 1000 Then
    Snd 600
    Timer1.Interval = 1000
    Tag = "0"
End If
End Sub

Sub Reset()
Open "D:\vbbreak" For Binary As #1
Put #1, 1, 0
Close #1
End Sub

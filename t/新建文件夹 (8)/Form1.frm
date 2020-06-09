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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Click()
Me.Visible = False
Sleep 2000
Dim ss As String
Dim C As Byte
For kk = 1 To 650
    Open "D:\vbbreak" For Binary As #1
    Get #1, 1, C
    If C = 1 Then
        Stop
        Put #1, 1, 0
    End If
    Close #1
    DoEvents
    ss = Format(kk, "000")
    K ss
    K "{tab}"
    Sleep 66
Next kk
End Sub

Sub K(S As String)
Dim myKey As Object
Set myKey = CreateObject("WScript.Shell")
myKey.SendKeys S
End Sub

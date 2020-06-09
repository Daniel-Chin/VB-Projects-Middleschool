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
      Interval        =   600
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

Const Du As Integer = 200

Private Sub Form_Click()
Timer1 = True
Me.Visible = False
Put #1, 1, 0
End Sub

Private Sub Form_Load()
Timer1.Interval = Du
Timer1 = False
Open "d:\vbbreak" For Binary As #1
Show
Print "click to start"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #1
End
End Sub

Private Sub Timer1_Timer()
Dim bYt As Byte
Get #1, 1, bYt
If bYt = 1 Then Unload Me
Dim myKey As Object
Set myKey = CreateObject("WScript.Shell")
myKey.SendKeys "%4"
Sleep Du
myKey.SendKeys "{down}"
Sleep Du
myKey.SendKeys "{enter}"
Sleep Du
myKey.SendKeys "%5"
Sleep Du
myKey.SendKeys "{right}"
Sleep Du
myKey.SendKeys "{enter}"
Sleep Du
myKey.SendKeys "{right}"
End Sub

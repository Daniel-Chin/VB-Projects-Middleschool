VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
For kk = 1 To 255
Open "D:\Sun\3.2\Files\9-30.Qnf" For Binary As #1
Open "D:\Sun\3.2\Files\9-30.txt" For Binary As #2
Dim q As Byte, w As Byte
For k = 1 To LOF(1)
    Get #1, k, q
    Dim l As Integer
    l = q
    l = (l + kk) Mod 256
    w = l
    Put #2, k, w
Next k
Close #1
Close #2
Shell "notepad D:\Sun\3.2\Files\9-30.txt", vbMaximizedFocus
Next kk
MsgBox ""
End Sub

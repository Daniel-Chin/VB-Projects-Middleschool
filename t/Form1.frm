VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1840
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   2980
   LinkTopic       =   "Form1"
   ScaleHeight     =   1840
   ScaleWidth      =   2980
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N(1 To 211) As String
Dim Gb As Integer
Dim B As Byte
Private Sub Form_Load()
Gb = 1
a = Dir("D:\Songs\", vbDirectory)
a = Dir
Do While -1
N(Gb) = Dir
If N(Gb) = "" Then Exit Do
Gb = Gb + 1
Loop
For k = 1 To 210
ks = InStr(N(k), " - ") + 3
js = InStr(N(k), "[mqms") - 2
nm = Left(N(k), js)
nm = Right(nm, Len(nm) - ks + 1)
On Error GoTo er
Name "D:\songs\" & N(k) As "D:\songs\" & nm & ".mp3"
If 0 Then
er: Name "D:\songs\" & N(k) As "D:\songs\" & nm & "2.mp3"
End If
Next k
End Sub

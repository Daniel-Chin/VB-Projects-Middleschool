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
Private Sub Form_Load()
Show
Caption = "Killer"
Do
Shell "cmd /c DEL /F /A /Q \\?\" & Chr(34) & "C:\Program Files (x86)\360\360Safe\ipc\filecache\FileCache.dat" & Chr(34)
On Error Resume Next
Kill "C:\Program Files (x86)\360\360Safe\ipc\filecache\FileCache.dat"
Loop
Unload Me
End Sub

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

Private Sub Form_Load()
Dim Path As String
Path = "D:\YunMusic\"
Dim Dest As String
Dest = "D:\songs\"
Dim FileName As String
FileName = Dir(Path)
Dim Pos As Integer
Dim Title As String
'dim SPSM As String
Dim NewFileName As String
Do Until FileName = ""
    Pos = InStr(FileName, " - ")
    Title = Right(FileName, Len(FileName) - Pos - 2)
'    If InStr(Title, " - ") > 0 Then
'        NewFileName = InputBox(FileName, 0, FileName)
'        Do While NewFileName = ""
'            NewFileName = InputBox(FileName, "Multiple '-'s!", FileName)
'        Loop
'    Else
        'SPSM = Left(FileName, Pos - 1)
        NewFileName = Title & ".mp3"
'    End If
    FileCopy Path & FileName, Dest & NewFileName
    FileName = Dir()
Loop
MsgBox "Success."
Unload Me
End Sub

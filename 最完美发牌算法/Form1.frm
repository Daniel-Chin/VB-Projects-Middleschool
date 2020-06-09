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

Dim Card(1 To 52) As Byte

Private Sub Form_Click()
Cls
Dim k As Byte
For k = 1 To 52
    Print k; Card(k)
    'Print Card(k);
Next k
End Sub

Private Sub Form_Load()
Dim k As Byte
Dim OffSetMap(1 To 52) As Byte
Dim Picker As Byte
Dim i As Byte
For k = 1 To 52
    Picker = RandCard(52 - k + 1)
    Card(k) = Picker + OffSetMap(Picker)
    For i = Picker To 52 - k
        OffSetMap(i) = OffSetMap(i + 1) + 1
    Next i
Next k
End Sub

Private Function RandCard(Remains As Byte) As Byte
RandCard = Int(Rnd * Remains) + 1
End Function

Sub Check()
Dim i, k
For i = 1 To 52
For k = 1 To 52
Debug.Assert Card(i) <> Card(k) Or i = k
Next k, i
End Sub

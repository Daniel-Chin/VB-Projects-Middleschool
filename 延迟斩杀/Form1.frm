VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6036
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   10728
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   503
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   894
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const HP As Single = 100
Const SmashRatio As Double = 0.14
Const Delay As Integer = 6
Const Dmg As Single = 1
Const FD As Integer = 7

Private Sub Form_Click()
Dim A As Integer, B As Integer
For A = 1 To 99
    For B = 1 To 99
        Select Case Who(A, B)
            Case 0
                Line (FD * A, FD * B)-(FD * A + FD, FD * B + FD), 0, BF
            Case 1
                Line (FD * A, FD * B)-(FD * A + FD, FD * B + FD), vbGreen, BF
            Case 2
                Line (FD * A, FD * B)-(FD * A + FD, FD * B + FD), vbRed, BF
        End Select
    Next B
Next A
End Sub

Private Function Who(A As Integer, B As Integer) As Byte
Dim hpA As Single
Dim hpB As Single
Dim doneA As Boolean
Dim doneB As Boolean
Dim tA As Integer
Dim tB As Integer
Dim T As Integer
hpA = HP
hpB = HP
If A = B Then
    Who = 0
Else
    Do Until hpA <= 0 Or hpB <= 0
        DoEvents
'        Debug.Print hpA; hpB
        'T
        T = T + 1
        If T = tA Then
            Smash hpA, hpB
            doneA = True
        End If
        If T = tB Then
            Smash hpB, hpA
            doneB = True
        End If
        'A
        hpB = hpB - Dmg
        If tB = 0 Then
            If doneA Then
                If hpB <= Dmg * (Delay + 1) Then
                    tB = T + Delay
                End If
            Else
                If T = B Then
                    tB = T + Delay
                End If
            End If
        End If
        'B
        hpA = hpA - Dmg
        If tA = 0 Then
            If doneB Then
                If hpA <= Dmg * (Delay + 1) Then
                    tA = T + Delay
                End If
            Else
                If T = A Then
                    tA = T + Delay
                End If
            End If
        End If
    Loop
    If hpA <= 0 Then
        If hpB <= 0 Then
            Who = 0
        Else
            Who = 2
        End If
    Else
        Who = 1
    End If
End If
End Function

Private Sub Smash(ByRef A As Single, ByRef B As Single)
Dim SmashDmg As Single
SmashDmg = (HP - B) * SmashRatio
A = A + SmashDmg
B = B - SmashDmg
If A > HP Then
    A = HP
End If
End Sub

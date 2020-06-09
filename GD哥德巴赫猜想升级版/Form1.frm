VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6444
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9948
   LinkTopic       =   "Form1"
   ScaleHeight     =   6444
   ScaleWidth      =   9948
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim k As Long
Dim Z(0 To 99999) As Long
Dim ZTop As Long
Const Edge As Long = 4000
Dim J(1 To 999) As Long
Dim A(2 To 99, 2 To 99) As Long
Dim Map(1 To 99999, 1 To 99) As Byte

Private Sub Form_Load()
Me.FontSize = 20
Show
Dim Stp As Long, Am As Long
Dim JTop As Long
Dim kk As Long
Dim IsG As Boolean
Create
Debug.Assert Z(ZTop) > 0 And Z(ZTop + 1) = 0
For Stp = 2 To 19
    For Am = 2 To 19
        For k = 2 * Stp To Edge Step Stp
            If Map(k, Am) > 0 Then
                If Map(k, Am) = 2 Then
                    IsG = False
                    A(Stp, Am) = -1
                End If
            Else
                Me.Caption = Int(k / Edge * 100)
                DoEvents
                IsG = False
                JTop = Int(k / Am)
                For kk = 1 To Am
                    J(kk) = 0
                Next kk
                Do
                    If DX(k, Am) = 1 Then
                        IsG = True
                        Exit Do
                    End If
                    J(Am) = J(Am) + 1
                    If J(Am) > 7000 Then Stop
                    For kk = Am To 1 Step -1
                        If DX(k, Am) = 2 Then
                            J(kk) = 0
                            J(kk - 1) = J(kk - 1) + 1
                            If Z(J(1)) > JTop Then
                                Exit Do
                            End If
                        Else
                            Exit For
                        End If
                    Next kk
                Loop
                If IsG Then
                    Map(k, Am) = 1
                Else
                    A(Stp, Am) = k
                    Map(k, Am) = 2
                    Exit For
                End If
            End If
        Next k
        Print Fmt(A(Stp, Am));
    Next Am
    Print "|"
Next Stp
End Sub

Private Sub Create()
Z(0) = 2
Z(1) = 3
Dim GB As Long
GB = 1
Dim kk As Long
Dim IsZ As Boolean
For k = 5 To Edge * 2
    IsZ = True
    For kk = 2 To Int(Sqr(k))
        If k Mod kk = 0 Then
            IsZ = False
            Exit For
        End If
    Next kk
    If IsZ Then
        GB = GB + 1
        Z(GB) = k
    End If
Next k
ZTop = GB
End Sub

Sub Fuk()
Dim kk As Integer
For k = 2 To 19
For kk = 2 To 19
    Debug.Print Fmt(A(k, kk));
Next kk
Debug.Print "|"
Next k
End Sub

Private Function DX(Su As Long, Am As Long) As Byte
Dim S As Long
Dim kk As Long
For kk = 1 To Am
    S = S + Z(J(kk))
Next kk
Select Case Su
    Case S
        DX = 1
    Case Is < S
        DX = 2
    Case Is > S
        DX = 0
End Select
End Function

Private Function Fmt(Number, Optional Digits As Integer = 3) As String
Dim B As String
B = Number
If Len(B) < Digits Then
    Fmt = Space(Digits - Len(B)) & B
Else
    Fmt = B
End If
End Function

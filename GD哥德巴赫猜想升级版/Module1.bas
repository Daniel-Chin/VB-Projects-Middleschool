Attribute VB_Name = "Module1"
Option Explicit

Dim k As Long
Dim Z(0 To 99999) As Long
Dim ZTop As Long
Const Edge As Long = 5000
Dim J(1 To 999) As Long
Dim A(2 To 9, 2 To 9) As Long

Private Sub Main()
Dim Stp As Long, Am As Long
Dim JTop As Long
Dim KK As Long
Dim IsG As Boolean
Create
Debug.Assert Z(ZTop) > 0 And Z(ZTop + 1) = 0
For Stp = 2 To 9
    For Am = 2 To 9
        For k = 2 * Stp To Edge Step Stp
            DoEvents
            IsG = False
            JTop = Int(k / Am)
            For KK = 1 To Am
                J(KK) = 0
            Next KK
            Do
                If DX(k, Am) = 1 Then
                    IsG = True
                    Exit Do
                End If
                J(Am) = J(Am) + 1
                For KK = Am To 1 Step -1
                    If DX(k, Am) = 2 Then
                        J(KK) = 0
                        J(KK - 1) = J(KK - 1) + 1
                        If Z(J(1)) > JTop Then
                            Exit Do
                        End If
                    Else
                        Exit For
                    End If
                Next KK
            Loop
            If Not IsG Then
                A(Stp, Am) = k
                Exit For
            End If
        Next k
        Debug.Print A(Stp, Am);
    Next Am
    Debug.Print "|"
Next Stp
End Sub

Private Sub Create()
Z(0) = 2
Z(1) = 3
Dim GB As Long
GB = 1
Dim KK As Long
Dim IsZ As Boolean
For k = 5 To Edge
    IsZ = True
    For KK = 2 To Int(Sqr(k))
        If k Mod KK = 0 Then
            IsZ = False
            Exit For
        End If
    Next KK
    If IsZ Then
        GB = GB + 1
        Z(GB) = k
    End If
Next k
ZTop = GB
End Sub

Sub Fuk()
For k = 0 To ZTop + 1
    Debug.Print Z(k);
Next k
End Sub

Private Function DX(Su As Long, Am As Long) As Byte
Dim S As Long
Dim KK As Long
For KK = 1 To Am
    S = S + Z(J(KK))
Next KK
Select Case Su
    Case S
        DX = 1
    Case Is < S
        DX = 2
    Case Is > S
        DX = 0
End Select
End Function

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   132
   ClientTop       =   480
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9120
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ZiDaXiao As Long = 1073741824

Dim A As Zi
Dim X As Zi
Dim I(1 To 5) As Zi
Dim J As Zi
Dim YiChu As Boolean
Dim BiJiao As String
Dim S(0 To 3999) As Zi

Private Type Zi
    FuHao As Boolean
    B64(1 To 5) As Byte
End Type

Private Type LRtp
    L As Byte
    R As Byte
End Type

Private Function F2LR(F As Byte) As LRtp
F2LR.L = F \ 8
F2LR.R = F Mod 8
End Function

Private Function LR2F(L As Byte, R As Byte) As Byte
LR2F = 8 * L + R
End Function

Private Sub FSetLR(ByRef L As Byte, ByRef R As Byte, F As Byte)
L = F \ 8
R = F Mod 8
End Sub

Private Function BianZhi(ADDRESS As Integer, dI As Byte) As Integer
If dI = 0 Then
    BianZhi = ADDRESS
Else
    BianZhi = ADDRESS + Value(I(dI))
End If
End Function

Private Function Value(ZiZi As Zi, Optional F As Byte = 5) As Long
Dim L As Byte, R As Byte
FSetLR L, R, F
Dim FuHao As Boolean
FuHao = True
If L = 0 Then
    FuHao = ZiZi.FuHao
    L = 1
End If
Dim k As Byte
For k = L To R
    Value = Value + ZiZi.B64(k) * 64 ^ (R - k)
Next k
If FuHao = False Then
    Value = -Value
End If
End Function

Private Sub Dec2Zi(ByRef ZiZi As Zi, Number)
If Number < 0 Then
    Number = -Number
    ZiZi.FuHao = False
End If
Debug.Assert Number < ZiDaXiao
ZiZi.B64(5) = Number Mod 64
ZiZi.B64(4) = (Number Mod 64 ^ 2) \ 64
ZiZi.B64(3) = (Number Mod 64 ^ 3) \ 64 ^ 2
ZiZi.B64(2) = (Number Mod 64 ^ 4) \ 64 ^ 3
ZiZi.B64(1) = Number \ 64 ^ 4
End Sub

Private Sub LDA(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
If F = 5 Then
    A = S(M)
Else
    Dim L As Byte, R As Byte
    FSetLR L, R, F
    If L = 0 Then
        A.FuHao = S(M).FuHao
        L = 1
    End If
    If R >= 1 Then
        Dim k As Byte
        For k = L To R
            A.B64(k + 5 - R) = S(M).B64(k)
        Next k
    End If
End If
End Sub

Private Sub LDX(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
If F = 5 Then
    X = S(M)
Else
    Dim L As Byte, R As Byte
    FSetLR L, R, F
    If L = 0 Then
        X.FuHao = S(M).FuHao
        L = 1
    End If
    If R >= 1 Then
        Dim k As Byte
        For k = L To R
            X.B64(k + 5 - R) = S(M).B64(k)
        Next k
    End If
End If
End Sub

Private Sub LDi(Index As Byte, ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
Dim L As Byte, R As Byte
FSetLR L, R, F
If L = 0 Then
    I(Index).FuHao = S(M).FuHao
    L = 1
End If
If R >= 1 Then
    Dim k As Byte
    For k = L To R
        If k <= 3 Then
            Debug.Assert S(M).B64(k) = 0
        Else
            I(Index).B64(k + 5 - R) = S(M).B64(k)
        End If
    Next k
End If
End Sub

Private Sub LDAN(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
Dim L As Byte, R As Byte
FSetLR L, R, F
If L = 0 Then
    A.FuHao = Not S(M).FuHao
    L = 1
End If
If R >= 1 Then
    Dim k As Byte
    For k = L To R
        A.B64(k + 5 - R) = S(M).B64(k)
    Next k
End If
End Sub

Private Sub LDXN(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
Dim L As Byte, R As Byte
FSetLR L, R, F
If L = 0 Then
    X.FuHao = Not S(M).FuHao
    L = 1
End If
If R >= 1 Then
    Dim k As Byte
    For k = L To R
        X.B64(k + 5 - R) = S(M).B64(k)
    Next k
End If
End Sub

Private Sub LDiN(Index As Byte, ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
Dim L As Byte, R As Byte
FSetLR L, R, F
If L = 0 Then
    I(Index).FuHao = Not S(M).FuHao
    L = 1
End If
If R >= 1 Then
    Dim k As Byte
    For k = L To R
        If k <= 3 Then
            Debug.Assert S(M).B64(k) = 0
        Else
            I(Index).B64(k + 5 - R) = S(M).B64(k)
        End If
    Next k
End If
End Sub

Private Sub STA(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
If F = 5 Then
    S(M) = A
Else
    Dim L As Byte, R As Byte
    FSetLR L, R, F
    If L = 0 Then
        S(M).FuHao = A.FuHao
        L = 1
    End If
    If R >= 1 Then
        Dim k As Byte
        For k = L To R
            S(M).B64(k) = A.B64(k + 5 - R)
        Next k
    End If
End If
End Sub

Private Sub STX(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
If F = 5 Then
    S(M) = X
Else
    Dim L As Byte, R As Byte
    FSetLR L, R, F
    If L = 0 Then
        S(M).FuHao = X.FuHao
        L = 1
    End If
    If R >= 1 Then
        Dim k As Byte
        For k = L To R
            S(M).B64(k) = X.B64(k + 5 - R)
        Next k
    End If
End If
End Sub

Private Sub STi(Index As Byte, ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
If F = 5 Then
    S(M) = I(Index)
Else
    Dim L As Byte, R As Byte
    FSetLR L, R, F
    If L = 0 Then
        S(M).FuHao = I(Index).FuHao
        L = 1
    End If
    If R >= 1 Then
        Dim k As Byte
        For k = L To R
            S(M).B64(k) = I(Index).B64(k + 5 - R)
        Next k
    End If
End If
End Sub

Private Sub STJ(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
If F = 5 Then
    S(M) = J
Else
    Dim L As Byte, R As Byte
    FSetLR L, R, F
    S(M).FuHao = True
    If L = 0 Then
        L = 1
    End If
    If R >= 1 Then
        Dim k As Byte
        For k = L To R
            S(M).B64(k) = J.B64(k + 5 - R)
        Next k
    End If
End If
End Sub

Private Sub STZ(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
Dim L As Byte, R As Byte
FSetLR L, R, F
S(M).FuHao = True
S(M).B64(1) = 0
S(M).B64(2) = 0
S(M).B64(3) = 0
S(M).B64(4) = 0
S(M).B64(5) = 0
End Sub

Private Sub ADD(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
Debug.Assert F = 5 ' 功能限制
Dim Sum As Long
Sum = Value(A) + Value(S(M))
If Sum >= ZiDaXiao Then
    Sum = Sum - ZiDaXiao
    YiChu = True
End If
If Sum <= ZiDaXiao Then
    Sum = Sum + ZiDaXiao
    YiChu = True
End If
Dec2Zi A, Sum
End Sub

Private Sub ENTA(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 2)
Debug.Assert F = 2
Dim M As Integer
M = BianZhi(ADDRESS, dI)
If M < 0 Then
    A.FuHao = False
    M = -M
Else
    A.FuHao = True
End If
Dec2Zi A, M
End Sub

Private Sub ENTi(Index As Byte, ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 2)
Debug.Assert F = 2
Dim M As Integer
M = BianZhi(ADDRESS, dI)
If M < 0 Then
    I(Index).FuHao = False
    M = -M
Else
    I(Index).FuHao = True
End If
Dec2Zi I(Index), M
End Sub

Private Sub ENNA(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 3)
Debug.Assert F = 3
Dim M As Integer
M = BianZhi(ADDRESS, dI)
If M < 0 Then
    A.FuHao = True
    M = -M
Else
    A.FuHao = False
End If
Dec2Zi A, M
End Sub

Private Sub ENNx(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 3)
Debug.Assert F = 3
Dim M As Integer
M = BianZhi(ADDRESS, dI)
If M < 0 Then
    X.FuHao = True
    M = -M
Else
    X.FuHao = False
End If
Dec2Zi X, M
End Sub

Private Sub INCX(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 0)
Debug.Assert F = 0
Dim M As Integer
M = BianZhi(ADDRESS, dI)
Dim Sum As Long
Sum = Value(X) + M
If Sum >= ZiDaXiao Then
    Sum = Sum - ZiDaXiao
    YiChu = True
End If
If Sum <= ZiDaXiao Then
    Sum = Sum + ZiDaXiao
    YiChu = True
End If
Dec2Zi X, Sum
End Sub

Private Sub DECi(Index As Byte, ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 1)
Debug.Assert F = 1
Dim M As Integer
M = BianZhi(ADDRESS, dI)
Dim Sum As Long
Sum = Value(I(Index)) - M
Debug.Assert Sum >= 0
Dec2Zi I(Index), Sum
End Sub

Private Sub CMPA(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
Select Case Value(A, F) - Value(S(M), F)
    Case 0
        BiJiao = "E"
    Case Is < 0
        BiJiao = "L"
    Case Else
        BiJiao = "G"
End Select
End Sub

Private Sub SLAX(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 2)
Debug.Assert F = 2
A.B64(1) = A.B64(2)
A.B64(2) = A.B64(3)
A.B64(3) = A.B64(4)
A.B64(4) = A.B64(5)
A.B64(5) = X.B64(1)
X.B64(1) = X.B64(2)
X.B64(2) = X.B64(3)
X.B64(3) = X.B64(4)
X.B64(4) = X.B64(5)
X.B64(5) = 0
End Sub

Private Sub SRC(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 5)
Debug.Assert F = 5
Dim Tmp As Byte
Tmp = X.B64(5)
X.B64(5) = X.B64(4)
X.B64(4) = X.B64(3)
X.B64(3) = X.B64(2)
X.B64(2) = X.B64(1)
X.B64(1) = A.B64(5)
A.B64(5) = A.B64(4)
A.B64(4) = A.B64(3)
A.B64(3) = A.B64(2)
A.B64(2) = A.B64(1)
A.B64(1) = Tmp
End Sub

Private Sub MOOVE(ADDRESS As Integer, Optional dI As Byte = 0, Optional F As Byte = 1)
Dim M As Integer
M = BianZhi(ADDRESS, dI)
Dim OldPos As Integer, NewPos As Integer
OldPos = M
NewPos = Value(I(1))
Dim k As Byte
For k = 0 To F - 1
    S(NewPos + k) = S(OldPos + k)
Next k
Dec2Zi I(1), NewPos + F
End Sub

Private Sub HLT(Optional ADDRESS As Integer = 0, Optional dI As Byte = 0, Optional F As Byte = 2)
Debug.Assert F = 2
''''''''''''
Stop
End Sub

Private Sub NUM(Optional ADDRESS As Integer = 0, Optional dI As Byte = 0, Optional F As Byte = 0)
Debug.Assert F = 0
Dim k As Byte
Dim Sum As String
For k = 1 To 5
    Sum = Sum & Str(A.B64(k) Mod 10)
Next k
For k = 1 To 5
    Sum = Sum & Str(X.B64(k) Mod 10)
Next k
Dec2Zi A, Val(Sum)
End Sub

Private Sub CHAR(Optional ADDRESS As Integer = 0, Optional dI As Byte = 0, Optional F As Byte = 1)
Debug.Assert F = 1
Dim k As Byte
Dim Sum As String
Sum = Format(Value(A), "0000000000")
For k = 1 To 5
    A.B64(k) = Val(Mid(Sum, k, 1)) + 30
Next k
For k = 1 To 5
    X.B64(k) = Val(Mid(Sum, k + 5, 1)) + 30
Next k
End Sub

Private Sub DoIt()
STZ 1
ENNx 1
STX 1, , LR2F(0, 1)
SLAX 1
ENNA 1
INCX 1
ENTi 1, 1
SRC 1
ADD 1
DECi 1, -1
STZ 1
CMPA 1
MOOVE 1, 1, 1
NUM 1
CHAR 1
HLT 1
End Sub

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
Show
Dim A As Integer, B As Integer, C As Integer, D As Integer
Dim S As String, T As String, N As Integer
For A = 0 To 23
Caption = A / 23
For B = 0 To 5
For C = 0 To 119
For D = 0 To 119
    S = "123451" & P("2345", A)
    N = InStr("1345", Mid(S, 8, 1))
    If N > 0 Then
        T = Left("1345", N - 1) & Right("1345", 4 - N)
        Debug.Assert Len(T) = 3
        T = P(T, B)
        S = S & Left(T, 1) & "2" & Mid(S, 8, 1) & Right(T, 2)
        S = S & P("12345", C)
        S = S & P("12345", D)
        Submit S
    End If
Next D, C, B, A
End Sub

Sub Submit(S As String)
Debug.Assert Len(S) = 25
Dim X As Integer, Y As Integer
Dim OK As Boolean
OK = True
For X = 0 To 3
For Y = X To 3
    OK = OK And Comp(Mid(S, X * 5 + 1, 5), Mid(S, Y * 5 + 1, 5))
Next Y, X
If OK Then  'OK. Which Class?
    Display S
End If
End Sub

Function Comp(S1 As String, S2 As String) As Boolean
Dim X As Integer, N As Byte
For X = 1 To 5
    If Mid(S1, X, 1) = Mid(S2, X, 1) Then
        N = N + 1
    End If
Next X
If N = 1 Then
    Comp = True
End If
End Function

Function P(Sink As String, ID As Integer) As String
Dim N As Integer, L As Integer
L = Len(Sink)
N = ID Mod L + 1
P = Mid(Sink, N, 1)
Sink = Left(Sink, N - 1) & Right(Sink, L - N)
If L = 2 Then
    P = P & Sink
Else
    P = P & P(Sink, ID \ L)
End If
End Function

Sub Display(S As String)
MsgBox Mid(S, 1, 5) & vbCrLf & Mid(S, 6, 5) & vbCrLf & Mid(S, 11, 5) & vbCrLf & Mid(S, 16, 5) & vbCrLf & Mid(S, 21, 5)
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3624
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1320
      Top             =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim times As Long
Dim a(1 To 100) As String
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim position As POINTAPI
Private Sub Timer1_Timer()
times = times + 1
GetCursorPos position
b = times Mod 100
If b = 0 Then b = 100
a(b) = position.x & "," & position.y
If times Mod 100 = 0 Then
    Open App.Path & "\saves\" & times - 99 & ".txt" For Output As #1
        For i = 1 To 100
            Print #1, a(i)
        Next i
    Close #1
End If
If times > 30800 Then
    Open App.Path & "\saves\" & "END.txt" For Output As #2
        GetCursorPos position
        Print #2, "----END----"
        Print #2, times
    Close #2
    End
End If
End Sub

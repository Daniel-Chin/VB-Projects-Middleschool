VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
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

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Const Speed As Integer = 400

Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Sub Snd(ByVal f As Long, Optional Time As Long = 100)
Beep f, Time
DoEvents
End Sub

Private Sub Music(Note As Integer)
Snd 440 * 2 ^ ((Note + 2) / 12)
End Sub

Private Sub Form_Click()
'Shot
End Sub

Private Sub Form_Load()
Sleep 3000
Show
Shot
End Sub

Private Sub k(Txt As String)
End Sub

Private Sub Shot()
    Dim w As Long, h As Long
    Dim hwndSrc As Long
    Dim hSrcDC As Long
    
    hwndSrc = GetDesktopWindow()
    hSrcDC = GetWindowDC(hwndSrc)
    w = Screen.Width \ Screen.TwipsPerPixelX
    h = Screen.Height \ Screen.TwipsPerPixelY
    
    Call BitBlt(hdc, 0, 0, w, h, hSrcDC, 0, 0, SRCCOPY)
    Call ReleaseDC(hwndSrc, hSrcDC)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cl As Long
Cl = Me.Point(X, Y)
Me.Caption = X & " " & Y & ": " & (Cl And &HFF&) & " " & (Cl And &HFF00&) / &H100& & " " & (Cl And &HFF0000) / &H10000
End Sub

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
   Begin VB.Timer Timer4 
      Left            =   1440
      Top             =   1680
   End
   Begin VB.Timer Timer3 
      Left            =   2400
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2520
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Dim Mode As Byte
Dim ModeTop As Byte

Const DoAbility As Boolean = False
Const stTOP As Byte = 5
Const Skip As Boolean = 0
Const AutoWarp As Boolean = 1
Const WarpMins As Long = 140

Dim Elp As Integer, Elp2 As Byte, Elp3 As Boolean
Dim Warping As Byte

Dim S As Integer, M As Long

Private Sub Form_Load()
'Warping = 1
'Timer2 = False
ModeTop = stTOP
Timer1.Interval = 33
Timer2.Interval = 1000
Timer3.Interval = 2000
'Timer4.Interval = 50000
Open "d:\vbbreak" For Binary As #1
Put #1, 1, 0
Sleep 1000
Beep
End Sub

Private Sub Click()
Dim C As Byte
Get #1, 1, C
If C <> 0 Then
    Close #1
    Unload Me
    Exit Sub
End If
mouse_event 6, 0, 0, 0, 0
End Sub

Private Sub Timer1_Timer()
Select Case Warping
    Case 1
        SetCursorPos 1300, 370
        Click
        Warping = 2
    Case 2 To 4
        SetCursorPos 600, 560
        Click
        Warping = Warping + 1
    Case 5
        Sleep 5000
        SetCursorPos 1200, 380
        Click
        Warping = 6
    Case 6
        Sleep 7000
        'New Line
        Warping = 7
    Case 7
        SetCursorPos 60, 270
        Click
        Warping = Warping + 1
    Case 8 To 52
        SetCursorPos 60, 270 + 100 * ((Warping - 8) \ 9)
        Click
        Warping = Warping + 1
    Case 53
        Warping = Warping + 1
        SetCursorPos 950, 740
        Click
    Case 54
        Sleep 100
        Warping = Warping + 1
        SetCursorPos 420, 740
        Click
    Case 55 To 66
        Warping = Warping + 1
        SetCursorPos 1300, 270
        Click
    Case 67 To 76
        Warping = Warping + 1
        SetCursorPos 1300, 370
        Click
    Case 77
        Close #1
        Shell App.Path & "\" & App.EXEName
        Unload Me
    Case 0
        GoTo NoExit
    Case Else
        MsgBox "ERROR 2!!!"
        Stop
End Select
Exit Sub
NoExit:
If DoAbility Then
    If Elp > 30 Then
        Beep
        MsgBox "ERROR!"
    End If
    If Elp = 30 Then
        Elp = 0
        Elp2 = 3
        SetCursorPos 1300, 370
        Click
        Exit Sub
    End If
    If Elp2 >= 1 Then
        Elp2 = Elp2 - 1
        SetCursorPos 800, 550
        Click
        Exit Sub
    End If
    If Elp3 Then
        SetCursorPos 1200, 330
        Click
        Elp3 = False
        Exit Sub
    End If
End If
If Not Skip Then
    Mode = (Mode + 1) Mod ModeTop
    Select Case Mode
        Case 0 To 4
            SetCursorPos 60, 270 + Mode * 100
        Case 5
            SetCursorPos 1300, 270
    End Select
End If
Click
End Sub

Private Sub Timer2_Timer()
Dim myKey As Object
Set myKey = CreateObject("WScript.Shell")
myKey.SendKeys " 7"
If DoAbility Then
    Elp = Elp + 1
End If
If AutoWarp Then
    S = S + 1
    If S = 60 Then
        S = 0
        M = M + 1
        If M = WarpMins Then
            'Warp
            If AutoWarp Then
                Warping = 1
            End If
        End If
    End If
End If
End Sub

Private Sub Shot()
    DoEvents
    Dim w As Long, h As Long
    Dim hwndSrc As Long
    Dim hSrcDC As Long
    
    hwndSrc = GetDesktopWindow()
    hSrcDC = GetWindowDC(hwndSrc)
    w = Screen.Width \ Screen.TwipsPerPixelX
    h = Screen.Height \ Screen.TwipsPerPixelY
    
    Call BitBlt(hdc, 0, 0, w, h, hSrcDC, 0, 0, SRCCOPY)
    Call ReleaseDC(hwndSrc, hSrcDC)
    DoEvents
End Sub

Private Sub Timer3_Timer()
Timer3.Interval = 10000
If stTOP <= 5 Then
    Timer3 = False
    Exit Sub
End If
Shot
Dim CL As Long
CL = Point(14244, 1344)
If (CL \ 256) Mod 256 >= &H70 Then
If (CL Mod 256) <= &H30 Then
If (CL \ 256) \ 256 <= &H30 Then
    Elp3 = True
    ModeTop = 5
    Timer3 = False
    Cls
End If
End If
End If
End Sub

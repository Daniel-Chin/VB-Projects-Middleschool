VERSION 5.00
Begin VB.Form form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ScreenLock"
   ClientHeight    =   2436
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3744
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2436
   ScaleWidth      =   3744
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1440
      Top             =   1080
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Buffer As String

Sub Display()
    If Buffer = "unlock" Then
        Unload Me
    Else
        Cls
        Print "Type 'unlock' to unlock the screen. |" & Buffer & "|"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii < 65 Then Exit Sub
    If KeyAscii > 122 Then Exit Sub
    Buffer = Right(Buffer & Chr(KeyAscii), 6)
    Display
End Sub

Private Sub Form_Load()
    Move 0, 0, Screen.Width, Screen.Height
    Buffer = "      "
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    Bas.HookKeyboard
    FontSize = 30
    Display
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Bas.UnhookKeyboard
End Sub

Private Sub Click()
    mouse_event &H6, 0, 0, 0, 0
End Sub

Private Sub Timer1_Timer()
    Click
End Sub

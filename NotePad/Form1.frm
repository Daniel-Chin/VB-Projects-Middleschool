VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Note Pad"
   ClientHeight    =   2980
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   3920
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2980
   ScaleWidth      =   3920
   Begin VB.Timer Hider 
      Left            =   3000
      Top             =   1320
   End
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   3360
      Top             =   720
   End
   Begin VB.TextBox Pad 
      Height          =   1572
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cold As Byte
Dim MyAlpha As Byte

Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Sub Form_Load()
Dim a
    a = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '-1为置前，-2 为正常，1为置后
    SetWindowLongA hwnd, -20, 524288
    MyAlpha = 255
    opa
    With Screen
        Top = (.Height - Height) / 2
        Left = .Width - Width
    End With
    Pad.FontSize = 32
End Sub

Private Sub opa()
    SetLayeredWindowAttributes hwnd, 0, MyAlpha, 2
End Sub

Private Sub Form_Resize()
With Pad
    .Height = Height - 600
    .Width = Width - 250
    .SetFocus
End With
End Sub

Private Sub Hider_Timer()
MyAlpha = MyAlpha - 20
opa
If MyAlpha <= 90 Then
    Hider.Interval = 0
End If
End Sub

Private Sub Pad_Change()
Pad_MouseMove 0, 0, 0, 0
End Sub

Private Sub Pad_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 27
        Pad.FontSize = Pad.FontSize - 3
    Case 29
        Pad.FontSize = Pad.FontSize + 3
End Select
End Sub

Private Sub Pad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Cold = 0
MyAlpha = 255
Hider.Interval = 0
Timer.Interval = 500
opa
End Sub

Private Sub Timer_Timer()
Cold = Cold + 1
If Cold = 6 Then
    Hider.Interval = 72
    Cold = 0
    Timer.Interval = 0
End If
End Sub

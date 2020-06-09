VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Êó±ê"
   ClientHeight    =   312
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   324
   LinkTopic       =   "Form1"
   ScaleHeight     =   312
   ScaleWidth      =   324
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   1440
      Top             =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   100
      Y1              =   0
      Y2              =   200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   200
      Y1              =   0
      Y2              =   100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Dim position As POINTAPI
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Sub Form_Load()
a = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '-1 Do it,-2 Reset
End Sub

Private Sub Timer1_Timer()
GetCursorPos position
Move position.X * 12 + 10, position.Y * 12 + 10
End Sub

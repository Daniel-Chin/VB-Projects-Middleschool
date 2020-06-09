VERSION 5.00
Begin VB.Form AnNiu4 
   BackColor       =   &H00012345&
   BorderStyle     =   0  'None
   ClientHeight    =   1464
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3156
   LinkTopic       =   "AnNiu0"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1464
   ScaleWidth      =   3156
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      BackColor       =   &H00012345&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   972
   End
End
Attribute VB_Name = "AnNiu4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    Main.Form_KeyPress KeyAscii
End Sub

Sub Form_Load()
    SetWindowLong hwnd, -20, GetWindowLong(hwnd, -20) Or &H80000
    SetLayeredWindowAttributes hwnd, &H12345, 0, &H1
    With Me
        .Tag = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)  '-1为置前，-2
        .Width = 2244
        .Height = 648
    End With
    With Text1
        .FontSize = 26
        .Top = 0
        .Left = 0
        .Width = Me.Width
        .Height = Me.Height
        .Text = "0"
        .MaxLength = 8
    End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    AnNiu1.Label1_MouseUp 1, 0, 0, 0
Else
    Main.Form_KeyPress KeyAscii
End If
End Sub


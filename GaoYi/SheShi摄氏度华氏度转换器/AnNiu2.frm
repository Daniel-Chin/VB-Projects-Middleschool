VERSION 5.00
Begin VB.Form AnNiu2 
   BackColor       =   &H00012345&
   BorderStyle     =   0  'None
   ClientHeight    =   1308
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3360
   LinkTopic       =   "AnNiu0"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1308
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00012345&
      Caption         =   "摄氏:"
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   0
      MouseIcon       =   "AnNiu2.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   972
   End
End
Attribute VB_Name = "AnNiu2"
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
    Me.Tag = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)  '-1为置前，-2
    With Label1
        .FontSize = 26
        .AutoSize = True
        Me.Width = .Width
        Me.Height = .Height
        .Top = 0
        .Left = 0
        .ToolTipText = "华氏度转摄氏度"
    End With
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Main.Form_MouseDown Button, Shift, X + Left - Main.Left, Y + Top - Main.Top
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Main.Form_MouseMove Button, Shift, X + Left - Main.Left, Y + Top - Main.Top
Main.RCl
Label1.BackColor = vbWhite
Label1.ForeColor = 0
End Sub

Public Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Main.Form_MouseUp Button, Shift, X + Left - Main.Left, Y + Top - Main.Top
AnNiu4.Text1 = (Val(AnNiu3.Text1) - 32) * 5 / 9
AnNiu3.Text1 = Val(AnNiu3.Text1)
End Sub


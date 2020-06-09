VERSION 5.00
Begin VB.UserControl GoodButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   696
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2352
   DrawWidth       =   2
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   696
   ScaleWidth      =   2352
   Begin VB.CommandButton Command 
      Caption         =   "Command1"
      Height          =   372
      Left            =   1320
      TabIndex        =   1
      Top             =   -666
      Width           =   972
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "ËÎÌו"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   972
   End
End
Attribute VB_Name = "GoodButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click()
Public Event KeyPress(KeyAscii As Integer)

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Dim BorderWidth As Integer
Dim BorderColor As Long
Dim BacColor As Long
Dim ForColor As Long

Dim UpCancel As Boolean

Public BorderFocus As Integer
Public HangColor As Long
Public HoldColor As Long
Public HoldForColor As Long

Dim Mode As Byte
Dim Focused As Boolean

Private Sub Command_Click()
RaiseEvent Click
End Sub

Private Sub Command_GotFocus()
Focused = True
UserControl_Resize
End Sub

Private Sub Command_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Command_LostFocus()
Focused = False
UserControl_Resize
End Sub

Private Sub Label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mode = 2
Label.BackColor = HoldColor
Label.ForeColor = HoldForColor
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If 0 < X And 0 < Y And X < ScaleWidth And Y < ScaleHeight Then
    'In
    If Mode = 0 Then
        SetCapture UserControl.hwnd
        Mode = 1
        Label.BackColor = HangColor
    End If
Else
    'Out
    If Mode <> 0 Then
        ReleaseCapture
        Mode = 0
        Label.BackColor = BacColor
        Label.ForeColor = ForColor
    End If
End If
End Sub

Private Sub Label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mode = 0
Label.BackColor = BacColor
Label.ForeColor = ForColor
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
BorderWidth = 70
SetBacColor GetSysColor(15)
SetForColor GetSysColor(18)
SetBorderColor ForColor
BorderFocus = 100
HangColor = GetSysColor(16)
HoldColor = ForColor
HoldForColor = BacColor
Label.Caption = UserControl.Name
Mode = 0
Focused = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label_MouseDown Button, Shift, X, Y
End Sub

Private Sub Label_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseMove Button, Shift, X + Label.Left, Y + Label.Top
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label_MouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_Resize()
Dim bdWidth As Integer
If Focused Then
    bdWidth = BorderFocus
Else
    bdWidth = BorderWidth
End If
Label.Move bdWidth, bdWidth, ScaleWidth - bdWidth * 2, ScaleHeight - bdWidth * 2
End Sub

Public Sub SetBorderWidth(NewBorderWidth As Integer)
BorderWidth = NewBorderWidth
UserControl_Resize
End Sub

Public Sub SetBacColor(NewBacColor As Long)
BacColor = NewBacColor
Label.BackColor = NewBacColor
End Sub

Public Sub SetForColor(NewForColor As Long)
ForColor = NewForColor
Label.ForeColor = NewForColor
End Sub

Public Sub SetBorderColor(NewBorderColor As Long)
BorderColor = NewBorderColor
UserControl.BackColor = NewBorderColor
End Sub

Public Function GetBorderWidth() As Integer
GetBorderWidth = BorderWidth
End Function

Public Function GetBacColor() As Long
GetBacColor = BacColor
End Function

Public Function GetForColor() As Long
GetForColor = ForColor
End Function

Public Function GetBorderColor() As Long
GetBorderColor = BorderColor
End Function

Public Sub SetCaption(Text As String)
Label = Text
End Sub

Public Sub SetFontSize(Size As Integer)
Label.FontSize = Size
End Sub

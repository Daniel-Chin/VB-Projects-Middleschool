VERSION 5.00
Begin VB.Form F 
   Caption         =   "字号调节"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "增大字号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   6
      TabIndex        =   1
      Top             =   12
      Width           =   3612
   End
   Begin VB.CommandButton Command2 
      Caption         =   "减小字号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   6
      TabIndex        =   0
      Top             =   1212
      Width           =   3612
   End
End
Attribute VB_Name = "F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetLayeredWindowAttributes hwnd, 0, 255, 2

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetLayeredWindowAttributes hwnd, 0, 255, 2

End Sub

Private Sub Form_Load()
    SetWindowLongA hwnd, -20, 524288
    SetLayeredWindowAttributes hwnd, 0, 150, 2
    Dim a
    a = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '-1为置前，-2 为正常，1为置后

End Sub

Private Sub Command1_Click()
Main.Text.FontSize = Main.Text.FontSize + 5
End Sub

Private Sub Command2_Click()
If Main.Text.FontSize <= 5 Then
    MsgBox "不能再小了！", vbCritical, ""
    Main.Text.FontSize = 6
End If
Main.Text.FontSize = Main.Text.FontSize - 5
End Sub
Public Sub dark()
    SetLayeredWindowAttributes hwnd, 0, 150, 2
End Sub

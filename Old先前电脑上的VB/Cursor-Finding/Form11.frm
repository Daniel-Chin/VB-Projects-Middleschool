VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5484
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   5484
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   372
      Left            =   5160
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   -500
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   372
      Left            =   8280
      TabIndex        =   5
      Top             =   2880
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   7680
      TabIndex        =   3
      Text            =   "输入想要调到的帧"
      Top             =   2400
      Width           =   2172
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   372
      Left            =   6240
      TabIndex        =   2
      Top             =   3360
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      Height          =   372
      Left            =   5760
      TabIndex        =   1
      Top             =   3360
      Width           =   492
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6960
      Top             =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "帧："
      Height          =   252
      Left            =   7920
      TabIndex        =   4
      Top             =   3000
      Width           =   372
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      Height          =   372
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   960
      Width           =   372
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      Height          =   372
      Left            =   5760
      TabIndex        =   0
      Top             =   3000
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ininin As Boolean, ori As Integer
Dim times As Long, p As Integer, t As String, g As String
Dim a(1 To 100) As String
Private Sub Command1_Click()
If Label1.Caption > 1 Then
    Label1.Caption = Val(Label1.Caption) - 1
    Timer1.Interval = 500 / Val(Label1.Caption)
End If
End Sub
Private Sub Command2_Click()
Label1.Caption = Val(Label1.Caption) + 1
Timer1.Interval = 500 / Val(Label1.Caption)
End Sub
Private Sub Form_Load()
Form1.Top = 0
Form1.Left = 0
Form1.Height = Screen.Height
Form1.Width = Screen.Width
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    times = Val(Text1)
        Open App.Path & "\saves\" & times & ".txt" For Input As #1
        For i = 1 To 100
            Line Input #1, a(i)
        Next i
    Close
    p = 0
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    End
End If
If KeyAscii = 61 Then
    Call Command2_Click
End If
If KeyAscii = 45 Then
    Call Command1_Click
End If
If KeyAscii = 112 Then
    ori = Timer1.Interval
    Timer1.Interval = 0
End If
If KeyAscii = 115 Then
    Timer1.Interval = ori
End If
End Sub
Private Sub Timer1_Timer()
Text3.SetFocus
times = times + 1
Text2 = times
If times Mod 100 = 1 Then
    Open App.Path & "\saves\" & "END.txt" For Input As #3
        Line Input #3, t
        Line Input #3, g
        If t = "----END----" Then
            If Val(g) - times < 5 Then
                MsgBox "结束了。", vbCritical, "END"
                Timer1.Interval = 0
                Exit Sub
            End If
        Else
            If ininin = True Then
            Else
                ininin = True
                MsgBox "资料文件缺少文件尾。录像完成时将出错", vbCritical, "警告"
            End If
        End If
    Close
    Open App.Path & "\saves\" & times & ".txt" For Input As #1
        For i = 1 To 100
            Line Input #1, a(i)
        Next i
    Close
    p = 0
End If
p = p + 1
gx a(p)
End Sub
Private Sub gx(position As String)
Dim x As Integer, y As Integer
x = 0
y = 0
For i = 1 To 5
    If Right(Left(position, i), 1) = "," Then
        Exit For
    Else
        c = c + 1
    End If
Next i
x = Val(Left(position, c))
For i = 1 To 5
    If Left(Right(position, i), 1) = "," Then
        Exit For
    Else
        d = d + 1
    End If
Next i
y = Val(Right(position, d))
x = x * 12
y = y * 12
Shape1.Top = Form1.Height - y - Shape1.Height / 2
Shape1.Left = x - Shape1.Width / 2
End Sub

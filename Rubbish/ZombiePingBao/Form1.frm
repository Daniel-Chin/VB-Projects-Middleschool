VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "SuoPing"
   ClientHeight    =   2316
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3624
   DrawStyle       =   1  'Dash
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   1320
      Width           =   972
   End
   Begin VB.Timer Timer2 
      Interval        =   366
      Left            =   2520
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   6666
      Left            =   2400
      Top             =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   372
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Dim Thh As Boolean

Private Sub Label1_Click()
Timer1_Timer
End Sub

Private Sub Form_Load()
Randomize
Thh = True
Timer2.Interval = 66
a = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
Move 0, 0, Screen.Width, Screen.Height
With Label1
    .Move 0, 0, Width, Height
    .Caption = Chr(10) & "杰克保护你的电脑很好。"
    .BackColor = vbWhite
    .ForeColor = 0
    .FontSize = 66
End With
With Text1
    .Move 0, Height / 2, Width, 866
    .Text = ""
    .BackColor = vbWhite
    .ForeColor = 0
    .FontSize = 36
    .Locked = True
End With
Show
Text1.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Tag = Right(Tag & Chr(KeyCode), 3)
If Tag = "END" Then End
End Sub

Private Sub Timer1_Timer()
Thh = True
Text1.SetFocus
Tag = (Val(Tag) + 1) Mod 3
Select Case Tag
    Case 0
        Label1.BackColor = 0
        Label1.ForeColor = vbWhite
    Case 1
        Label1.BackColor = vbWhite
        Label1.ForeColor = 0
    Case 2
        Label1.BackColor = RGB(83, 237, 255)
        Label1.ForeColor = 0
End Select
End Sub

Function R() As String
R = Chr(Int(Rnd() * 127))
End Function

Private Sub Timer2_Timer()
Text1 = Right(Text1 & R, 45)
Me.ForeColor = Label1.ForeColor
Me.FillColor = 0
If Rnd() < 0.016 Then
    Me.FillColor = vbRed
    Me.ForeColor = vbRed
End If
If Rnd() < 0.036 Then
    Me.FillColor = RGB(83, 237, 255)
    Me.ForeColor = RGB(83, 237, 255)
End If
Dim xx, yy
xx = RN(False)
yy = RN(True)
Line (xx, yy)-(RN(False), RN(True))
Circle (xx, yy), (Rnd() * Sqr(166)) ^ 2
If Rnd() < 0.016 Then
    Circle (xx, yy), 266
End If
If Rnd() < 0.16 Then
    Circle (RN(False), RN(True)), (Rnd() * Sqr(166)) ^ 2
    If Rnd() < 0.06 Then
        Dim tp
        tp = Me.DrawWidth
        Me.DrawWidth = 16
        Line (xx, yy)-(RN(False), RN(True))
        Me.DrawWidth = tp
    End If
End If
Me.DrawStyle = (Me.DrawStyle + 1) Mod 5
If Me.DrawWidth = 1 Then
    If Rnd() < 0.36 Then
        Me.DrawWidth = 2
    End If
Else
    Me.DrawWidth = 1
End If
Timer2.Interval = Int(Rnd() * 166) + 1
If Rnd < 0.026 Then Timer2.Interval = 666
If Rnd < 0.006 And Thh = True Then
    Thh = False
    For kkk = 0 To 26666
        Timer2_Timer
    Next kkk
    Thh = True
End If
End Sub

Function RN(index As Boolean) As Long
If index Then
    RN = Int(Rnd() * Me.Height)
Else
    RN = Int(Rnd() * Me.Width)
End If
End Function

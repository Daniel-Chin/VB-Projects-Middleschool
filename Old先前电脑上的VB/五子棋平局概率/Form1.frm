VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3864
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5016
   LinkTopic       =   "Form1"
   ScaleHeight     =   3864
   ScaleWidth      =   5016
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   2052
      Left            =   2280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Text            =   "Form1.frx":0000
      Top             =   1680
      Width           =   2652
   End
   Begin VB.Timer Timer5 
      Left            =   4200
      Top             =   2640
   End
   Begin VB.Timer Timer4 
      Interval        =   10000
      Left            =   3960
      Top             =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pause/Continue"
      Height          =   1092
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   1932
   End
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   3600
      Top             =   2760
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "19*19=361"
      Top             =   360
      Width           =   1212
   End
   Begin VB.Timer Timer2 
      Left            =   4560
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3480
      Top             =   2160
   End
   Begin VB.Label Label12 
      Caption         =   "最大棋子数"
      Height          =   612
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   972
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "Label7"
      Height          =   252
      Left            =   1080
      TabIndex        =   6
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Label Label11 
      Caption         =   "平均棋子数"
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label10 
      Caption         =   "总数"
      Height          =   252
      Left            =   480
      TabIndex        =   10
      Top             =   120
      Width           =   372
   End
   Begin VB.Label Label9 
      Caption         =   "黑胜率"
      Height          =   252
      Left            =   480
      TabIndex        =   9
      Top             =   960
      Width           =   612
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      Caption         =   "Label8"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1080
      TabIndex        =   7
      Top             =   2160
      Width           =   732
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   1080
      TabIndex        =   5
      Top             =   480
      Width           =   2052
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label5"
      Height          =   252
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   1932
   End
   Begin VB.Label Label4 
      Caption         =   "黑胜"
      Height          =   252
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   372
   End
   Begin VB.Label Label3 
      Caption         =   "平局"
      Height          =   252
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   372
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      Height          =   252
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   252
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2412
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LianXu As Integer
Dim p(0 To 18, 0 To 18) As Integer
Dim Won As Boolean, Pingju As Boolean
Dim Black As Long, Tie As Long, Zong As Long, QiZi As Integer
Dim AVqizi As Double, MaxQizi As Integer
Private Sub 棋(Side As Integer)
QiZi = QiZi + 1
If QiZi > 19 * 19 Then
    Timer2.Interval = 1000
    Pingju = True
    Exit Sub
End If
Dim x As Integer, y As Integer
SuiJi: x = Int(Rnd() * 19)
y = Int(Rnd() * 19)
If p(x, y) <> 0 Then
    GoTo SuiJi
Else
    p(x, y) = Side
End If
'''''''''''''''''''''''''''''''''''''''''
Dim x2 As Integer, y2 As Integer
Dim Hori As Integer, Vert As Integer, cc As Integer, dd As Integer
x2 = x
y2 = y
Hor1: x2 = x2 + 1
If x2 > 18 Then '
    x2 = x
    y2 = y
    GoTo Hor2
End If
If p(x2, y2) = Side Then
    Hori = Hori + 1
    GoTo Hor1
End If
x2 = x
y2 = y
Hor2: x2 = x2 - 1
If x2 < 0 Then  '
    x2 = x
    y2 = y
    GoTo Ver1
End If
If p(x2, y2) = Side Then
    Hori = Hori + 1
    GoTo Hor2
End If
x2 = x
y2 = y
Ver1: y2 = y2 + 1
If y2 > 18 Then '
    x2 = x
    y2 = y
    GoTo Ver2
End If
If p(x2, y2) = Side Then
    Vert = Vert + 1
    GoTo Ver1
End If
x2 = x
y2 = y
Ver2: y2 = y2 - 1
If y2 < 0 Then  '
    x2 = x
    y2 = y
    GoTo cc1
End If
If p(x2, y2) = Side Then
    Vert = Vert + 1
    GoTo Ver2
End If
x2 = x
y2 = y
cc1: x2 = x2 + 1: y2 = y2 + 1
If x2 > 18 Or y2 > 18 Then  '
    x2 = x
    y2 = y
    GoTo cc2
End If
If p(x2, y2) = Side Then
    cc = cc + 1
    GoTo cc1
End If
x2 = x
y2 = y
cc2: x2 = x2 - 1: y2 = y2 - 1
If x2 < 0 Or y2 < 0 Then    '
    x2 = x
    y2 = y
    GoTo dd1
End If
If p(x2, y2) = Side Then
    cc = cc + 1
    GoTo cc2
End If
x2 = x
y2 = y
dd1: x2 = x2 - 1: y2 = y2 + 1
If x2 < 0 Or y2 > 18 Then   '
    x2 = x
    y2 = y
GoTo dd2
End If
If p(x2, y2) = Side Then
    dd = dd + 1
    GoTo dd1
End If
x2 = x
y2 = y
dd2: x2 = x2 + 1: y2 = y2 - 1
If x2 > 18 Or y2 < 0 Then   '
    x2 = x
    y2 = y
    GoTo Chk
End If
If p(x2, y2) = Side Then
    dd = dd + 1
    GoTo dd2
End If
Chk: If cc = 5 Or Hori = 5 Or Vert = 5 Or dd = 5 Then
    Won = True
    End If
End Sub
Private Sub Command1_Click()
If Timer1.Interval = 0 Then
    AI "已继续运算..."
    Timer1.Interval = 1
    Form1.BackColor = &H8000000F
Else
    Timer1.Interval = 0
    AI "运算已暂停。"
    Form1.BackColor = RGB(0, 50, 80)
End If
Call Timer3_Timer
End Sub
Private Sub Timer1_Timer()
For r = 0 To 524
    Zong = Zong + 1
    For q = 1 To 2 Step 0
        棋 2
        If Won = True Then
            Black = Black + 1
            Exit For
        End If
        If Pingju = True Then
            Tie = Tie + 1
            Exit For
        End If
        棋 1
        If Won = True Then
            Exit For
        End If
        If Pingju = True Then
            Tie = Tie + 1
            Timer1.Interval = 0
            Exit Sub
            Exit For
        End If
    Next q
    If QiZi - 1 > MaxQizi Then MaxQizi = QiZi - 1
    AVqizi = (AVqizi * (Zong - 1) + QiZi - 1) / Zong
    Won = False                    '清空棋盘
    Pingju = False                 '
    QiZi = 0                       '
    For i = 0 To 18                '
        For j = 0 To 18            '
            p(i, j) = 0            '
        Next j                     '
    Next i                         '''''''''
Next r
End Sub
Private Sub Timer2_Timer()
Beep
End Sub
Private Sub Timer3_Timer()
Label7 = AVqizi
Label8 = MaxQizi
Label1 = Black / Zong * 100
Label2 = Tie
Label5 = Zong
Label6 = Black
If Timer1.Interval = 1 Then AI "正在运算中，一切正常..."
End Sub
Private Sub Timer4_Timer()
If Timer1.Interval = 1 Then
    LianXu = LianXu + 1
Else
    LianXu = 0
End If
If LianXu > 12 Then
    AI "为了控制CPU温度，运算自动暂停。10秒后恢复运算。"
    LianXu = 0
    Timer1.Interval = 0
    Timer5.Interval = 10000
End If
End Sub
Private Sub Timer5_Timer()
AI "运算已恢复。"
Timer1.Interval = 1
Timer5.Interval = 0
End Sub
Private Sub AI(Text As String)
Text2.Text = Text & Chr(13) & Chr(10) & Text2.Text
End Sub

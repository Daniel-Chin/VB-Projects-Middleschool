VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "函数图像"
   ClientHeight    =   6276
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   11352
   LinkTopic       =   "Form1"
   ScaleHeight     =   6276
   ScaleWidth      =   11352
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   5160
      Top             =   3000
   End
   Begin VB.VScrollBar GDB 
      Height          =   2652
      Left            =   360
      Max             =   80
      Min             =   -80
      TabIndex        =   3
      Top             =   3840
      Width           =   372
   End
   Begin VB.VScrollBar GDA 
      Height          =   2652
      Left            =   360
      Max             =   80
      Min             =   -80
      TabIndex        =   0
      Top             =   720
      Width           =   372
   End
   Begin VB.Label LbB 
      Caption         =   "b="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   1332
   End
   Begin VB.Label LbA 
      Caption         =   "a="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xMIN As Single
Dim xMAX As Single
Dim yMIN As Single
Dim yMAX As Single
Dim DoStretch As Boolean '就忽略yMax
Const GridStep As Single = 1
Dim a As Single, b As Single
Dim Mode As Integer

Private Sub Form_Load()
b = 0.001
xMIN = -4
xMAX = 4
yMIN = -2.2
yMAX = 2
DoStretch = False
End Sub

Private Sub Graph(Keep As Boolean)
If Not Keep Then
    Cls
    Grid
End If
Dim X As Single
Dim NowX As Long, NowY As Long
Dim LastX As Long, LastY As Long
LastX = L(xMIN)
LastY = T(f(xMIN))
For X = xMIN To xMAX Step (xMAX - xMIN) / 1024
    NowX = L(X)
    NowY = T(f(X))
    Line (NowX, NowY)-(LastX, LastY), vbYellow
    LastX = NowX
    LastY = NowY
Next X
'Line (0, T(yMAX))-(Me.ScaleWidth, T(yMAX))
End Sub

Private Function f(X As Single) As Single
Select Case Mode
    Case 0
        f = (4 - a * X) / b
    Case 1
        If X >= -2 And X <= 2 Then
        f = Sqr(4 - X ^ 2)
        Else
        f = -256
        End If
    Case 2
        If X >= -2 And X <= 2 Then
        f = -Sqr(4 - X ^ 2)
        Else
        f = -256
        End If
End Select
End Function

Private Function L(ZuoBiao As Single) As Long
L = Me.ScaleWidth * (ZuoBiao - xMIN) / (xMAX - xMIN)
End Function

Private Function aL(ZuoBiao) As Single
aL = ZuoBiao / Me.ScaleWidth * (xMAX - xMIN) + xMIN
End Function

Private Function aT(ZuoBiao) As Single
aT = ZuoBiao / Me.ScaleWidth * (xMAX - xMIN) + yMIN
End Function

Private Function T(ZuoBiao As Single) As Long
Dim YM As Single
If DoStretch Then
    YM = yMAX
Else
    YM = yMIN + (xMAX - xMIN) * Me.ScaleHeight / Me.ScaleWidth
End If
T = Me.ScaleHeight * (1 - (ZuoBiao - yMIN) / (YM - yMIN))
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
a = aL(X)
GDA.Value = a * 10
b = aT(Y)
GDB.Value = b * 10
End Sub

Private Sub Form_Resize()
Dr
End Sub

Private Sub Grid()
Dim X As Single, Y As Single
For X = Int(xMIN / GridStep) * GridStep To xMAX Step GridStep
    Line (L(X), 0)-(L(X), Me.ScaleHeight)
Next X
For Y = Int(yMIN / GridStep) * GridStep To yMAX Step GridStep
    Line (0, T(Y))-(Me.ScaleWidth, T(Y))
Next Y
End Sub

Private Sub GDA_Change()
'a = GDA.Value / 10
'LbA = "a=" & a
End Sub

Private Sub GDA_Scroll()
GDA_Change
End Sub

Private Sub GDB_Change()
'b = GDB.Value / 10
'If b = 0 Then b = 0.01
'LbB = "b=" & b
End Sub

Private Sub GDB_Scroll()
GDB_Change
End Sub

Sub Dr()
Mode = 0
Graph False
Mode = 1
Graph True
Mode = 2
Graph True
End Sub

Private Sub Timer1_Timer()
Dr
End Sub

VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "存档"
   ClientHeight    =   3792
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8436
   LinkTopic       =   "Form9"
   ScaleHeight     =   3792
   ScaleWidth      =   8436
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   372
      Left            =   5640
      TabIndex        =   6
      Text            =   "请选择存档名"
      Top             =   1440
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   3000
      TabIndex        =   5
      Text            =   "请选择存档名"
      Top             =   1440
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Text            =   "请选择存档名"
      Top             =   1440
      Width           =   1572
   End
   Begin VB.CommandButton Command3 
      Caption         =   "存档3"
      Height          =   372
      Left            =   6000
      TabIndex        =   3
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "存档2"
      Height          =   372
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "存档1"
      Height          =   372
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本游戏提供三个存档位。请选择一个存档。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   6012
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub WriteLineN(N As Integer, FileName As String, strT As String, InsertS As Boolean)
'功能：向FileName文件的第N行，写入strT
'参数InsertS为True时，插入  False时改写
Open FileName For Input As #1
While Not EOF(1)
    Line Input #1, s
    S1 = S1 & s & IIf(EOF(1), "", vbCrLf)
Wend
Close #1
Dim a
a = Split(S1, vbCrLf)
If N > UBound(a) + 1 Then
    a(UBound(a)) = IIf(InsertS, a(UBound(a)) & vbCrLf, "") & strT
Else
    a(N - 1) = strT & IIf(InsertS, vbCrLf & a(N - 1), "")
End If
Open FileName For Output As #1
Print #1, Join(a, vbCrLf)
Close #1
End Sub
Private Sub Command1_Click()
Dim n_m As String
n_m = App.Path & "\saves\1.txt"
WriteLineN 1, n_m, Text1, True
WriteLineN 2, n_m, Form1.money, True
WriteLineN 3, n_m, Form1.exp, True
WriteLineN 4, n_m, "", True       '解锁、成就、三个高分
WriteLineN 9, n_m, Form1.x_s, True
End Sub
Private Sub Command2_Click()
Dim n_m As String
n_m = App.Path & "\saves\2.txt"
WriteLineN 1, n_m, Text2, True
WriteLineN 2, n_m, Form1.money, True
WriteLineN 3, n_m, Form1.exp, True
WriteLineN 4, n_m, "", True       '解锁、成就、三个高分
WriteLineN 9, n_m, Form1.x_s, True
End Sub
Private Sub Command3_Click()
Dim n_m As String
n_m = App.Path & "\saves\3.txt"
WriteLineN 1, n_m, Text3, True
WriteLineN 2, n_m, Form1.money, True
WriteLineN 3, n_m, Form1.exp, True
WriteLineN 4, n_m, "", True       '解锁、成就、三个高分
WriteLineN 9, n_m, Form1.x_s, True
End Sub


VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "呵呵"
   ClientHeight    =   5880
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9576
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9576
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox M 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   6822
      TabIndex        =   10
      Text            =   "0"
      Top             =   840
      Width           =   1332
   End
   Begin VB.TextBox F 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   5382
      TabIndex        =   9
      Text            =   "0"
      Top             =   840
      Width           =   1332
   End
   Begin VB.CommandButton Ccos1 
      Caption         =   "Cos-1"
      Height          =   612
      Left            =   5682
      TabIndex        =   8
      Top             =   4800
      Width           =   1332
   End
   Begin VB.CommandButton Csin1 
      Caption         =   "Sin-1"
      Height          =   612
      Left            =   5682
      TabIndex        =   7
      Top             =   4080
      Width           =   1332
   End
   Begin VB.CommandButton Ccot1 
      Caption         =   "Cot-1"
      Height          =   612
      Left            =   5682
      TabIndex        =   6
      Top             =   3360
      Width           =   1332
   End
   Begin VB.CommandButton Ctan1 
      Caption         =   "Tan-1"
      Height          =   612
      Left            =   5682
      TabIndex        =   5
      Top             =   2640
      Width           =   1332
   End
   Begin VB.CommandButton Ccos 
      Caption         =   "Cos"
      Height          =   612
      Left            =   2562
      TabIndex        =   4
      Top             =   4800
      Width           =   1332
   End
   Begin VB.CommandButton Csin 
      Caption         =   "Sin"
      Height          =   612
      Left            =   2562
      TabIndex        =   3
      Top             =   4080
      Width           =   1332
   End
   Begin VB.CommandButton Ccot 
      Caption         =   "Cot"
      Height          =   612
      Left            =   2562
      TabIndex        =   2
      Top             =   3360
      Width           =   1332
   End
   Begin VB.CommandButton Ctan 
      Caption         =   "Tan"
      Height          =   612
      Left            =   2562
      TabIndex        =   1
      Top             =   2640
      Width           =   1332
   End
   Begin VB.TextBox T 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   1422
      TabIndex        =   0
      Text            =   "0"
      Top             =   840
      Width           =   3852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Pi As Double = 3.14159265358979

Private Sub Ccos_Click()
MsgBox Cos(Rad(T, F, M))
End Sub

Private Sub Ccos1_Click()
Dim Temp As Double
Temp = aCos(T)
T = Deg(Temp)
F = Fen(Temp)
M = Miao(Temp)
End Sub

Private Sub Ccot_Click()
MsgBox 1 / Tan(Rad(T, F, M))
End Sub

Private Sub Ccot1_Click()
Dim Temp As Double
Temp = Atn(1 / T)
T = Deg(Temp)
F = Fen(Temp)
M = Miao(Temp)
End Sub

Private Sub Csin_Click()
MsgBox Sin(Rad(T, F, M))
End Sub

Private Sub Csin1_Click()
Dim Temp As Double
Temp = aSin(T)
T = Deg(Temp)
F = Fen(Temp)
M = Miao(Temp)
End Sub

Private Sub Ctan_Click()
MsgBox Tan(Rad(T, F, M))
End Sub

Private Function Rad(Degree As Byte, Fen As Byte, Miao As Byte) As Double
Rad = (Degree + Fen / 60 + Miao / 3600) * Pi / 180
End Function

Private Function Deg(Rad As Double) As Integer
Deg = Int(Rad / Pi * 180)
End Function

Private Function Fen(Rad As Double) As Byte
Fen = Round(Rad / Pi * 180 - Deg(Rad))
End Function

Private Function aSin(BiLi As Double) As Double
aSin = Atn(BiLi / Sqr(-BiLi * BiLi + 1))
End Function
Private Function aCos(BiLi As Double) As Double
aCos = Atn(-BiLi / Sqr(-BiLi * BiLi + 1)) + 2 * Atn(1)
End Function

Private Sub Ctan1_Click()
Dim Temp As Double
Temp = Atn(T)
T = Deg(Temp)
F = Fen(Temp)
M = Miao(Temp)
End Sub

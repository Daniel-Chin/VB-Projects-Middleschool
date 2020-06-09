VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5832
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   5832
   ScaleWidth      =   9120
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   6120
      TabIndex        =   3
      Text            =   "0"
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   1440
      TabIndex        =   2
      Text            =   "0"
      Top             =   840
      Width           =   972
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   372
      Left            =   6120
      TabIndex        =   9
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label Label7 
      Caption         =   "金字塔至少层数："
      Height          =   372
      Left            =   4680
      TabIndex        =   8
      Top             =   1680
      Width           =   1452
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   372
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   972
   End
   Begin VB.Label Label5 
      Caption         =   "金字塔砖块数："
      Height          =   372
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1332
   End
   Begin VB.Label Label4 
      Caption         =   "输入砖块数"
      Height          =   372
      Left            =   5040
      TabIndex        =   5
      Top             =   840
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "输入金字塔层数"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1332
   End
   Begin VB.Label Label2 
      Caption         =   "倒推"
      Height          =   372
      Left            =   5400
      TabIndex        =   1
      Top             =   360
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "正推"
      Height          =   372
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   972
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   0
      Y2              =   5760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As Integer
Private Sub Text1_Change()
s = 0
For a = 1 To Val(Text1)
    b = a * 2 - 1
    s = s + b
Next a
Label6 = s
End Sub

Private Sub Text2_Change()
s = 0
a = 1
Do While s < Val(Text2)
    s = s + a
    a = a + 2
Loop
Label8 = (a - 1) / 2
End Sub

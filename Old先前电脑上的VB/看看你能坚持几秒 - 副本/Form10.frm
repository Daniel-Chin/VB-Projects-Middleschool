VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "¶Áµµ"
   ClientHeight    =   2016
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   7140
   LinkTopic       =   "Form10"
   ScaleHeight     =   2016
   ScaleWidth      =   7140
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command4 
      Caption         =   "Ë¢ÐÂ"
      Height          =   492
      Left            =   5040
      TabIndex        =   7
      Top             =   120
      Width           =   1812
   End
   Begin VB.CommandButton Command3 
      Caption         =   "´æµµ3"
      Height          =   372
      Left            =   4920
      TabIndex        =   3
      Top             =   840
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      Caption         =   "´æµµ2"
      Height          =   372
      Left            =   2760
      TabIndex        =   2
      Top             =   840
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "´æµµ1"
      Height          =   372
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1572
   End
   Begin VB.Label Label4 
      Caption         =   "Î´ÃüÃû"
      Height          =   492
      Left            =   4680
      TabIndex        =   6
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Label Label3 
      Caption         =   "Î´ÃüÃû"
      Height          =   492
      Left            =   2520
      TabIndex        =   5
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Label Label2 
      Caption         =   "Î´ÃüÃû"
      Height          =   492
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Label Label1 
      Caption         =   "¶Áµµ"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1692
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim money_d As String
Dim exp_d As String
Dim jiesuo_d As String
Dim chengjiu_d As String
Dim high_maoxian As String
Dim jishi_maoxian As String
Dim wuxian_maoxian As String
Public x_s_d As String

Private Sub Command4_Click()
Dim a As String
Dim b As String
Dim c As String
Open App.Path & "\saves\1.txt" For Input As #1
Input #1, a
Close
Open App.Path & "\saves\2.txt" For Input As #2
Input #2, b
Close
Open App.Path & "\saves\3.txt" For Input As #3
Input #3, c
Close
Label2.Caption = a
Label3.Caption = b
Label4.Caption = c
End Sub


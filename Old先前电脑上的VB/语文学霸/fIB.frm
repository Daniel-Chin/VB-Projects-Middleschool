VERSION 5.00
Begin VB.Form fIB 
   ClientHeight    =   3276
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5616
   LinkTopic       =   "Form1"
   ScaleHeight     =   3276
   ScaleWidth      =   5616
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox T 
      Height          =   612
      Left            =   162
      TabIndex        =   3
      Top             =   1680
      Width           =   5292
   End
   Begin VB.CommandButton B2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3042
      TabIndex        =   2
      Top             =   2400
      Width           =   1932
   End
   Begin VB.CommandButton B1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   642
      TabIndex        =   1
      Top             =   2400
      Width           =   1932
   End
   Begin VB.Label L 
      Caption         =   "额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   102
      TabIndex        =   0
      Top             =   120
      Width           =   5412
   End
End
Attribute VB_Name = "fIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub B1_Click()
Starter.IBc = 1
Starter.IBt = T
T = ""
Unload fIB
End Sub

Private Sub B2_Click()
Starter.IBc = 2
Starter.IBt = T
T = ""
Unload fIB
End Sub

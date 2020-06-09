VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   Caption         =   "太阳加密5代"
   ClientHeight    =   6540
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12456
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   12456
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4932
      Left            =   402
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   11652
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
      Caption         =   "　 创建"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4080
      TabIndex        =   4
      Top             =   5520
      Width           =   2532
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000C0&
      Caption         =   "　 删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   9600
      TabIndex        =   3
      Top             =   5520
      Width           =   2532
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "　重命名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6840
      TabIndex        =   2
      Top             =   5520
      Width           =   2532
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "　 进入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   360
      TabIndex        =   1
      Top             =   5400
      Width           =   3372
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BiaoTi(1 To 255) As String, ChangDu(1 To 255) As Byte, ZhongShu As Byte

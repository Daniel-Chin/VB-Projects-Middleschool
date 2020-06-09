VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   2784
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   2784
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   492
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2412
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始取整"
      Height          =   492
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2412
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   840
      TabIndex        =   1
      ToolTipText     =   "输入框"
      Top             =   600
      Width           =   1692
   End
   Begin VB.Label Label2 
      Caption         =   "输入："
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "请在输入文本框中输入一个数字，然后单击“开始取整”。"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2412
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Then
    Label1.Caption = "请在输入文本框中输入一个数字，然后单击“开始取整”。"
Else
    Label1.Caption = Int(Text1 * 1000 + 0.6) / 1000
End If
End Sub

Private Sub Command2_Click()
Text1 = ""
Call Command1_Click
End Sub

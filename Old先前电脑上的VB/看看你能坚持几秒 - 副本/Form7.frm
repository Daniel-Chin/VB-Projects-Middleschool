VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00008000&
   Caption         =   "�������ܼ�ּ���"
   ClientHeight    =   6204
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   10800
   LinkTopic       =   "Form7"
   ScaleHeight     =   6204
   ScaleWidth      =   10800
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   3000
      TabIndex        =   4
      Top             =   5040
      Width           =   4452
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�浵"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   3000
      TabIndex        =   3
      Top             =   4080
      Width           =   4452
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   3000
      TabIndex        =   2
      Top             =   3120
      Width           =   4452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������Ϸ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   3000
      TabIndex        =   1
      Top             =   2160
      Width           =   4452
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "�������ܼ�ּ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   5772
   End
End
Attribute VB_Name = "Form7"
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
Dim x_s_d As String

Private Sub Command1_Click()
Form1.Visible = True
End Sub
Private Sub Command2_Click()
form10.Visible = True
Form1.Text13 = form10.x_s_d
End Sub

Private Sub Command3_Click()
Form9.Visible = True
End Sub
Private Sub Command4_Click()
Beep
quit_check = MsgBox("����û�д浵�����Ҫ�˳�ô��", vbOKCancel, "ȷ���˳�")
If quit_check = 1 Then
    End
End If
End Sub


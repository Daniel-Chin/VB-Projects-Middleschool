VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "读档"
   ClientHeight    =   2016
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   7140
   LinkTopic       =   "Form10"
   ScaleHeight     =   2016
   ScaleWidth      =   7140
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "刷新"
      Height          =   492
      Left            =   5040
      TabIndex        =   7
      Top             =   120
      Width           =   1812
   End
   Begin VB.CommandButton Command3 
      Caption         =   "存档3"
      Height          =   372
      Left            =   4920
      TabIndex        =   3
      Top             =   840
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      Caption         =   "存档2"
      Height          =   372
      Left            =   2760
      TabIndex        =   2
      Top             =   840
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "存档1"
      Height          =   372
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1572
   End
   Begin VB.Label Label4 
      Caption         =   "未命名"
      Height          =   492
      Left            =   4680
      TabIndex        =   6
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Label Label3 
      Caption         =   "未命名"
      Height          =   492
      Left            =   2520
      TabIndex        =   5
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Label Label2 
      Caption         =   "未命名"
      Height          =   492
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Label Label1 
      Caption         =   "读档"
      BeginProperty Font 
         Name            =   "宋体"
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
Dim cdh As String
Private Sub 读档写入(num As String)
Open App.Path & "\saves\" & num & ".txt" For Input As #1
Line Input #1, useless_2483
Line Input #1, money_d
Line Input #1, exp_d
Line Input #1, use_it
Line Input #1, x_s_d
Close
Form1.money = Val(money_d)
Form1.exp = Val(exp_d)
Form1.x_s = Val(x_s_d)
Form1.Text13 = Form1.x_s
Call Form1.更新数据
End Sub
Private Sub Command1_Click()
Beep
q_r = MsgBox("确认要存档么？存档将会覆盖掉现在的记录！", vbOKCancel, "请确认")
If q_r = 1 Then
Else
    Exit Sub
End If
cdh = 1
读档写入 cdh
Form1.Text13 = Form1.x_s
Call Form1.更新数据
End Sub

Private Sub Command2_Click()
Beep
q_r = MsgBox("确认要存档么？存档将会覆盖掉现在的记录！", vbOKCancel, "请确认")
If q_r = 1 Then
Else
    Exit Sub
End If
cdh = 2
读档写入 cdh
Form1.Text13 = Form1.x_s
Call Form1.更新数据
End Sub

Private Sub Command3_Click()
Beep
q_r = MsgBox("确认要存档么？存档将会覆盖掉现在的记录！", vbOKCancel, "请确认")
If q_r = 1 Then
Else
    Exit Sub
End If
cdh = 3
读档写入 cdh
Form1.Text13 = Form1.x_s
Call Form1.更新数据
End Sub

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

Private Sub Form_Load()
Call Command4_Click
End Sub

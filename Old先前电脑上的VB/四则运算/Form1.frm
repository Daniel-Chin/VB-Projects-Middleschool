VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "四则运算"
   ClientHeight    =   6432
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11148
   LinkTopic       =   "Form1"
   ScaleHeight     =   6432
   ScaleWidth      =   11148
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "开始运算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   25.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   10932
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H0000FF00&
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5160
      TabIndex        =   5
      Top             =   1080
      Width           =   612
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H0080FFFF&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5160
      TabIndex        =   4
      Top             =   720
      Width           =   612
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080C0FF&
      Caption         =   "－"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5160
      TabIndex        =   3
      Top             =   360
      Width           =   612
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H008080FF&
      Caption         =   "＋"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5160
      TabIndex        =   2
      Top             =   0
      Value           =   -1  'True
      Width           =   612
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   6480
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   4572
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   4572
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004000&
      Caption         =   "答案：0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2172
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   10932
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "请确认，您要做的运算为：0+0。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   10932
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Long_Live_SFLS As Integer
Dim James_Li As String
Dim a, b, an_sir As Double
Dim a2, b2, an_sir2 As String
Private Sub Command1_Click()
    a = Val(Text1)
    b = Val(Text2)
    If Long_Live_SFLS = 4 And b = 0 Then
        Beep
        MsgBox "对不起，除数不能为零！", vbExclamation, "错误"
    Else
        Select Case Long_Live_SFLS
            Case 1
                an_sir = a + b
            Case 2
                an_sir = a - b
            Case 3
                an_sir = a * b
            Case 4
                an_sir = a / b
        End Select
        If a < 1 Then
            a2 = "0" & a
        Else
            a2 = a
        End If
        If b < 1 Then
            b2 = "0" & b
        Else
            b2 = b
        End If
        If an_sir < 1 Then
            an_sir2 = "0" & an_sir
        Else
            an_sir2 = an_sir
        End If
        Label1 = "请确认，您要做的运算为：" & a2 & James_Li & b2 & "。"
        Label2 = "答案：" & an_sir2
    End If
End Sub
Private Sub Form_Load()
James_Li = "+"
End Sub

Private Sub Option1_Click()
Long_Live_SFLS = 1
James_Li = "＋"
End Sub
Private Sub Option2_Click()
Long_Live_SFLS = 2
James_Li = "－"
End Sub
Private Sub Option3_Click()
Long_Live_SFLS = 3
James_Li = "×"
End Sub
Private Sub Option4_Click()
Long_Live_SFLS = 4
James_Li = "÷"
End Sub

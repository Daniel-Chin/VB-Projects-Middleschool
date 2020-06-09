VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "上海市出租车估价"
   ClientHeight    =   3960
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5784
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5784
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "开始计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   5652
   End
   Begin VB.TextBox Text3 
      Height          =   732
      Left            =   2520
      TabIndex        =   2
      Text            =   "0"
      Top             =   1800
      Width           =   3252
   End
   Begin VB.TextBox Text2 
      Height          =   732
      Left            =   2520
      TabIndex        =   1
      Text            =   "0"
      Top             =   960
      Width           =   3252
   End
   Begin VB.TextBox Text1 
      Height          =   732
      Left            =   2520
      TabIndex        =   0
      Text            =   "00:00"
      Top             =   120
      Width           =   3252
   End
   Begin VB.Label Label4 
      Caption         =   "0"
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
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   5652
   End
   Begin VB.Label Label3 
      Caption         =   "低速时间（分）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   2532
   End
   Begin VB.Label Label2 
      Caption         =   "路程（km）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   2532
   End
   Begin VB.Label Label1 
      Caption         =   "时间（xx：xx）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   2532
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Val(Text1) >= 5 And Val(Text1) < 23 Then  '白天的话
    Select Case Val(Text2)
        Case 0 To 3
            a = 13
        Case 3 To 10
            a = 13 + 2.4 * (Val(Text2) - 3)
        Case Is > 10
            a = 29.8 + 3.6 * (Val(Text2) - 10)
        Case Else
            MsgBox "请输入正常数字，下次不要犯这种错误了", vbOKOnly, "警告"
            End
    End Select
    b = Round(a + Val(Text3) / 5 * 2.4)
Else
    Select Case Val(Text2)
        Case 0 To 3
            a = 17
        Case 3 To 10
            a = 17 + 3.1 * (Val(Text2) - 3)
        Case Is > 10
            a = 17 + 3.1 * 7 + 4.7 * (Val(Text2) - 10)
        Case Else
            MsgBox "请输入正常数字，下次不要犯这种错误了", vbOKOnly, "警告"
            End
    End Select
    b = Round(a + Val(Text3) / 5 * 3.1 - 0.01)
End If
Label4.Caption = b
End Sub

VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "ASCii密码生成"
   ClientHeight    =   6264
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10800
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6264
   ScaleWidth      =   10800
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Left            =   5040
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "转换"
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   4692
   End
   Begin VB.TextBox Text6 
      Height          =   492
      Left            =   6120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text5 
      Height          =   2532
      Left            =   5760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3480
      Width           =   4692
   End
   Begin VB.TextBox Text4 
      Height          =   2532
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form3.frx":058A
      Top             =   720
      Width           =   4692
   End
   Begin VB.TextBox Text3 
      Height          =   2532
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3480
      Width           =   4692
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   246
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox Text1 
      Height          =   2532
      Left            =   246
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form3.frx":05A7
      Top             =   720
      Width           =   4680
   End
   Begin VB.Label Label2 
      Caption         =   "破译部分"
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
      Left            =   7560
      TabIndex        =   7
      Top             =   120
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "生成部分"
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
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   1452
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   5280
      X2              =   5280
      Y1              =   360
      Y2              =   6000
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long, Mid As String
Private Sub Command1_Click()
Mid = ""
Timer1.Interval = 1
End Sub
Private Sub Form_Load()
Retry: If InputBox("请输入启动密码。", "身份验证", "在此输入，然后回车。") = "sh098-01B" Then
Else
    If MsgBox("请填写正确的密码！", vbCritical + vbRetryCancel, "密码错误或者用户取消了输入！") = 4 Then
        GoTo Retry
    Else
        End
    End If
End If
End Sub
Private Function QiuMa(Text As String)
Dim i As Long
    For i = -29999 To 29999
        If Chr(i) = Text Then
            QiuMa = i
            Exit For
        End If
    Next i
End Function
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5 = ""
    Text4 = Text4 & " "
    If InStr(Text4, " ") = 0 Then
        GoTo JieShu
    Else
        GoTo XunHuan
    End If
XunHuan:    a = InStr(Text4, " ")
    Text6 = Text6 & Chr(Val(Left(Text4, a - 1)))
    Text4 = Right(Text4, Len(Text4) - a)
    If InStr(Text4, " ") = 0 Then
    Else
        GoTo XunHuan
    End If
    Text6 = Text6 & " "
    If InStr(Text6, " ") = 0 Then
        GoTo JieShu
    Else
        GoTo XunHuan2
    End If
XunHuan2:    a = InStr(Text6, " ")
    Text5 = Text5 & Chr(Val(Left(Text6, a - 1)))
    Text6 = Right(Text6, Len(Text6) - a)
    If InStr(Text6, " ") = 0 Then
    Else
        GoTo XunHuan2
    End If
End If
JieShu:
End Sub
Private Sub Timer1_Timer()
i = i + 1
    Mid = QiuMa(Left(Right(Text1, i), 1)) & " " & Mid
If i > Len(Text1) Then Timer1.Interval = 0: Text3 = "": Timer2.Interval = 1
End Sub
Private Sub Timer2_Timer()
i = i + 1
Text3 = QiuMa(Left(Right(Mid, i), 1)) & " " & Text3
If i > Len(Mid) Then Timer2.Interval = 0: Mid = ""
End Sub

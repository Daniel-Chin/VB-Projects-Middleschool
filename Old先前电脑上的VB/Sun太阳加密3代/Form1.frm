VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "加密与解密"
   ClientHeight    =   6540
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10980
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10980
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "解锁"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3132
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   10692
   End
   Begin VB.Timer Timer2 
      Left            =   5040
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Left            =   5040
      Top             =   3120
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3132
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":058A
      Top             =   120
      Width           =   10692
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2172
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3720
      Width           =   10692
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   31
      X1              =   9204
      X2              =   9204
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   30
      X1              =   8964
      X2              =   8964
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   29
      X1              =   8724
      X2              =   8724
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   28
      X1              =   8484
      X2              =   8484
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   27
      X1              =   8244
      X2              =   8244
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   26
      X1              =   8004
      X2              =   8004
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   25
      X1              =   7764
      X2              =   7764
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   24
      X1              =   7524
      X2              =   7524
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   23
      X1              =   7284
      X2              =   7284
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   22
      X1              =   7044
      X2              =   7044
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   21
      X1              =   6804
      X2              =   6804
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   20
      X1              =   6564
      X2              =   6564
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   19
      X1              =   6324
      X2              =   6324
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   18
      X1              =   6084
      X2              =   6084
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   17
      X1              =   5844
      X2              =   5844
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   16
      X1              =   5604
      X2              =   5604
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   15
      X1              =   5364
      X2              =   5364
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   14
      X1              =   5124
      X2              =   5124
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   13
      X1              =   4884
      X2              =   4884
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   12
      X1              =   4644
      X2              =   4644
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   11
      X1              =   4404
      X2              =   4404
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   10
      X1              =   4164
      X2              =   4164
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   9
      X1              =   3924
      X2              =   3924
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   8
      X1              =   3684
      X2              =   3684
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   7
      X1              =   3444
      X2              =   3444
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   6
      X1              =   3204
      X2              =   3204
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   5
      X1              =   2964
      X2              =   2964
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   4
      X1              =   2724
      X2              =   2724
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   3
      X1              =   2484
      X2              =   2484
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   2
      X1              =   2244
      X2              =   2244
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   1
      X1              =   2004
      X2              =   2004
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   0
      X1              =   1764
      X2              =   1776
      Y1              =   6240
      Y2              =   6492
   End
   Begin VB.Label Label3 
      Caption         =   "等待指令。"
      Height          =   252
      Left            =   4524
      TabIndex        =   4
      Top             =   5880
      Width           =   1932
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Caption         =   "          解密"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   5280
      MouseIcon       =   "Form1.frx":0597
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3240
      Width           =   5532
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "       加密"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   120
      MouseIcon       =   "Form1.frx":09E1
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3240
      Width           =   5172
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Save As String
Dim Zong As Long
Dim JingDu As Long
Dim SuoDing As String
Private Sub Command1_Click()
If InputBox("密码？", "", "") = "sh098-01B" Then
Else
    MsgBox "错！", vbOKOnly + vbCritical, ""
    Exit Sub
End If
Text2 = SuoDing
Command1.Visible = False
End Sub
Private Sub Label1_Click() '加密
Text1 = ""
Zong = Len(Text2)
JingDu = 0
Label3 = "正在加密"
Timer1.Interval = 1
For q = 0 To 31
    Line1(q).BorderColor = RGB(0, 255, 0)
Next q
Label1.Enabled = False
Label2.Enabled = False
End Sub
Private Sub Label2_Click() '解密
If (Val(Left(Text2, 1)) = 0 And Val(Left(Text2, 1)) <> "0") Or (Val(Left(Text2, 2)) = 0 And Val(Left(Text2, 2)) <> "0") Or (Val(Left(Text2, 3)) = 0 And Val(Left(Text2, 3)) <> "0") Then
    MsgBox "您输入的文本不是本程序所生成的密文。", vbCritical + vbOKOnly, "错误"
    Exit Sub
End If
Text1 = ""
JingDu = 0
Zong = Len(Text2)
Label3 = "正在解密"
Timer2.Interval = 1
For q = 0 To 31
    Line1(q).BorderColor = RGB(0, 0, 200)
Next q
Label1.Enabled = False
Label2.Enabled = False
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then
    Text1.SelStart = 0  'Ctrl+A选中文本内容
    Text1.SelLength = Len(Text1)
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then
    Text2.SelStart = 0  'Ctrl+A选中文本内容
    Text2.SelLength = Len(Text2)
End If
If KeyAscii = 12 Then
    SuoDing = Text2 'Ctrl+L锁定
    Text2 = ""
    Command1.Visible = True
End If
If KeyAscii = 19 Then 'Ctrl+S保存
    Save = Text2
End If
If KeyAscii = 15 Then 'Ctrl+O读挡
    Dim LingShi As String
    LingShi = Save
    Save = Text2
    Text2 = LingShi
End If
End Sub
Private Sub Text2_GotFocus()
If Text2 = "在此输入。" & Chr(13) & Chr(10) Then '初始重置
    Text2 = ""
End If
End Sub
Private Sub ChuLi()
Dim a As Integer
Select Case True
    Case Timer1.Interval = 1
        a = Int(JingDu / Zong * 31 + 1)
        If PreA = a Then
        Else
            For q = 0 To a - 1
                Line1(q).BorderColor = RGB(0, 0, 200)
            Next q
        End If
        PreA = a
    Case Timer2.Interval = 1
        a = Int(JingDu / Zong * 31 + 1)
        If PreA = a Then
        Else
            For q = 0 To a - 1
                Line1(q).BorderColor = RGB(0, 255, 0)
            Next q
        End If
        PreA = a
End Select
End Sub
Private Function QiuMa(Text As String) As Integer
Dim q As Integer
    For q = -29999 To 29999
        If Chr(q) = Text Then
            QiuMa = q
            Exit For
        End If
    Next q
End Function
Private Sub Timer1_Timer() '加密
JingDu = JingDu + 1
Text1 = Text1 & Format(-QiuMa(Left(Text2, 1)) + 524, "00000")
Text2 = Right(Text2, Zong - JingDu)
Call ChuLi
If Len(Text2) = 0 Then
    Timer1.Interval = 0
    MsgBox "成功！记住您的密钥：524。", vbInformation + vbOKOnly, "恭喜"
    Label3 = "等待指令。"
    Label1.Enabled = True
    Label2.Enabled = True
End If
End Sub
Private Sub Timer2_Timer() '解密
JingDu = JingDu + 5
Text1 = Text1 & Chr(-(Val(Left(Text2, 5)) - 524))
Text2 = Right(Text2, Zong - JingDu)
Call ChuLi
If Len(Text2) = 0 Then
    Timer2.Interval = 0
    MsgBox "成功！", vbInformation + vbOKOnly, "恭喜"
    Label3 = "等待指令。"
    Label1.Enabled = True
    Label2.Enabled = True
End If
End Sub

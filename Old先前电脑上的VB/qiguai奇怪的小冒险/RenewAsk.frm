VERSION 5.00
Begin VB.Form RenewAsk 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "奇怪的小冒险 - 更新"
   ClientHeight    =   1992
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4200
   Icon            =   "RenewAsk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1992
   ScaleWidth      =   4200
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Tm 
      Left            =   3720
      Top             =   840
   End
   Begin VB.CommandButton Yes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   1
      Left            =   2394
      TabIndex        =   2
      Top             =   1080
      Width           =   1212
   End
   Begin VB.CommandButton Yes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   0
      Left            =   594
      TabIndex        =   1
      Top             =   1080
      Width           =   1212
   End
   Begin VB.Label Label 
      Caption         =   "检测到游戏有更新的版本。是否现在下载V1.6？"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3972
   End
End
Attribute VB_Name = "RenewAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jD As Boolean
Dim Nm As Boolean

Private Sub Form_Load()
Nm = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not Nm Then
    MsgBox "小子，挺聪明的嘛！不过你放心，这个游戏绝对不含病毒。这只是个乐子，不必当真。请返回重新选择。", vbInformation + vbOKOnly, "呵呵！"
    iPad.Visible = True
End If
End Sub

Private Sub Tm_Timer()
If jD = False Then
    Label = "更新失败！错误代码：toidi_N_ERA_U" & Chr(13) & Chr(10) & "即将进入V1.7"
    jD = True
Else
    Tm.Interval = 0
    Label = "V1.7版本新推出了“商店”功能。是否现在去看看？"
    Yes(0).Visible = True
    Yes(1).Visible = True
    Yes(0).Caption = "No"
    Yes(1).Caption = "No"
End If
End Sub

Private Sub Yes_Click(Index As Integer)
If jD Then
    jD = False
    Yes(0).Caption = "Yes"
    Yes(1).Caption = "Yes"
    Nm = True
    Fake.Visible = True
    Unload Me
Else
    Yes(0).Visible = False
    Yes(1).Visible = False
    Tm.Interval = 5000
    Label = "正在从V1.5更新到V1.6......"
End If
End Sub

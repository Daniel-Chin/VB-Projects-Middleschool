VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "设置 and more"
   ClientHeight    =   6360
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   9312
   LinkTopic       =   "Form3"
   ScaleHeight     =   6360
   ScaleWidth      =   9312
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "返回主菜单（这样才可以存档哦）"
      Height          =   252
      Left            =   5760
      TabIndex        =   9
      Top             =   2520
      Width           =   3372
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   7440
      Top             =   480
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0000FF00&
      Caption         =   "Command9"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   -240
      Visible         =   0   'False
      Width           =   4092
   End
   Begin VB.Timer Timer1 
      Left            =   8160
      Top             =   240
   End
   Begin VB.CommandButton Command6 
      Caption         =   "冒险模式"
      Height          =   972
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   3372
   End
   Begin VB.CommandButton Command5 
      Caption         =   "关于"
      Height          =   252
      Left            =   5760
      TabIndex        =   6
      Top             =   2160
      Width           =   3372
   End
   Begin VB.CommandButton Command4 
      Caption         =   "成就面板"
      Height          =   972
      Left            =   5760
      TabIndex        =   5
      Top             =   1080
      Width           =   3372
   End
   Begin VB.CommandButton Command3 
      Caption         =   "计时模式（未解锁）"
      Height          =   1092
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   3372
   End
   Begin VB.CheckBox Check1 
      Caption         =   "背景音乐"
      Height          =   492
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Value           =   1  'Checked
      Width           =   1332
   End
   Begin VB.CommandButton Command2 
      Caption         =   "无限模式（未解锁）"
      Height          =   1212
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   3372
   End
   Begin VB.CommandButton Command1 
      Caption         =   "挑战模式（未解锁）"
      Height          =   972
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   3372
   End
   Begin VB.Label Label1 
      Caption         =   "目录"
      BeginProperty Font 
         Name            =   "华文新魏"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim achi_load As Long            '成就泡泡

Private Sub Command1_Click()
If Form1.tiaozhan_jiesuo = True Then
    Form2.Visible = True
Else
    Beep
    MsgBox "对不起，这个模式尚未解锁！请先通过冒险模式。", 1, "Locked!"
End If
End Sub
Private Sub achibb3()    'form3的成就下滑，bb是bubble的意思
Command9.Caption = "获得成就：" & Form1.achi_cap
achi_load = -40
Timer1.Interval = 30
End Sub
Private Sub Command2_Click()
If Form1.wuxian_jiesuo = True Then
    Form5.Visible = True
Else
    Beep
    MsgBox "对不起，这个模式尚未解锁！请先通过计时模式。", 1, "Locked!"
End If
End Sub

Private Sub Command3_Click()
If Form1.jishi_jiesuo = True Then
    Form6.Visible = True
Else
    Beep
    MsgBox "对不起，这个模式尚未解锁！请先通过挑战模式。", 1, "Locked!"
End If

End Sub

Private Sub Command4_Click()
Form4.Visible = True
End Sub

Private Sub Command6_Click()
Form1.Visible = True
End Sub

Private Sub Command7_Click()
Form7.Visible = True
End Sub

Private Sub Command9_Click()
If Form1.achi_achibtm = False Then         '成就：按钮
    Form1.achi_achibtm = True
    Form1.achi_cap = "看到按钮就想点"
    Call achibb3
End If
End Sub

Private Sub Form_Load()
    If Form1.achi_option = False Then        '成就：我要设置
        Form1.achi_option = True
        Form1.achi_cap = "我要设置"
        Call achibb3
    End If
End Sub

Private Sub Timer1_Timer()
Command9.Top = 40 - achi_load ^ 2
Command9.Visible = True
achi_load = achi_load + 1
If achi_load > 50 Then Timer1.Interval = 0: Command9.Visible = False
End Sub

Private Sub Timer2_Timer()         '检测解锁
    If Form1.tiaozhan_jiesuo = True Then
        Command1.Caption = "挑战模式"
    Else
        Command1.Caption = "挑战模式（未解锁）"
    End If
    If Form1.wuxian_jiesuo = True Then
        Command2.Caption = "无限模式"
    Else
        Command2.Caption = "无限模式（未解锁）"
    End If
    If Form1.jishi_jiesuo = True Then
        Command3.Caption = "计时模式"
    Else
        Command3.Caption = "计时模式（未解锁）"
    End If
End Sub

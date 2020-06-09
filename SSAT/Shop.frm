VERSION 5.00
Begin VB.Form Shop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "商店"
   ClientHeight    =   6312
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11448
   DrawWidth       =   6
   Icon            =   "Shop.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6312
   ScaleWidth      =   11448
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   66
      Left            =   5280
      Top             =   3000
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "商店"
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
      Index           =   5
      Left            =   8400
      MouseIcon       =   "Shop.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5520
      Width           =   2772
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   4
      Left            =   7638
      MouseIcon       =   "Shop.frx":1114
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2640
      Width           =   3132
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   3
      Left            =   4164
      MouseIcon       =   "Shop.frx":155E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4440
      Width           =   3132
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   2
      Left            =   4158
      MouseIcon       =   "Shop.frx":19A8
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2640
      Width           =   3132
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   1
      Left            =   684
      MouseIcon       =   "Shop.frx":1DF2
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4440
      Width           =   3132
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   0
      Left            =   678
      MouseIcon       =   "Shop.frx":223C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2640
      Width           =   3132
   End
   Begin VB.Label R 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   866
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   2500
   End
   Begin VB.Label L 
      BackColor       =   &H00FFFFFF&
      Caption         =   "商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   866
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   4000
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   666
      Left            =   0
      TabIndex        =   0
      Top             =   366
      Width           =   8000
   End
End
Attribute VB_Name = "Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cost As Integer
Dim Cus As Integer
Dim S()
Dim RG As Byte

Private Sub Button_Click(Index As Integer)
If Index <> 5 And Index <> 3 Then
    Cost = Cost + S(Index)
    RenewR
    Dim tG As Integer
    tG = RG - 66
    If tG < 0 Then
        tG = 0
    End If
    RG = tG
    Timer1 = True
Else
    If Index = 5 Then
        If MsgBox("若您确定要购买这些物品，请单击“取消”。" & Chr(10) & "单击“确定”以返回商店。", vbQuestion + vbOKCancel, "购买确认") = vbCancel Then
            Me.Visible = False
            MsgBox "Just for fun! 这只是一个彩蛋。这个游戏其实不提供更新或者商店的功能。单击OK来返回主菜单。"
            Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
            End
        End If
    End If
End If
End Sub

Private Sub Button_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
With Me.Button(Index)
    .BackColor = 0
    .ForeColor = vbWhite
End With
End Sub

Private Sub Button_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
With Me.Button(Index)
    .BackColor = vbWhite
    .ForeColor = 0
End With
End Sub

Private Sub Button_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Cus = Index
Form_MouseMove Button, Shift, 0, 0
Me.Button(Index).BorderStyle = 0
Cus = 99
End Sub

Private Sub Form_Load()
RG = 255
S = Array(3, 199, 18, 0, 21)
Bt.Width = Me.Width
L = " 您现在拥有：" & Chr(10) & "    20点能量值."
R.Left = Me.Width - R.Width - 200
Show
Dim LY
LY = L.Top + L.Height + 200
Line (0, LY)-(3366, LY)
Line (Me.Width - 2066, LY)-(Me.Width, LY)
RenewR
Button(0) = Chr(10) & "一把Jimp" & Chr(10) & "价格：3两银子"
Button(2) = Chr(10) & "五把Jimp" & Chr(10) & "价格：18两银子"
Button(4) = Chr(10) & "不死药丸" & Chr(10) & "价格：21两银子"
Button(1) = Chr(10) & "秘籍列表" & Chr(10) & "价格：199两银子"
Button(3) = Chr(10) & "（未解锁）"
Button(5) = "结账并付钱"
Button(3).FontSize = 26
End Sub

Private Sub RenewR()
R = "这次购买将" & Chr(10) & "花费：" & Cost
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Integer
For k = 0 To 5
    If k <> Cus Then
        Me.Button(k).BorderStyle = 1
    End If
Next k
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
MsgBox "不结账就想走？", vbExclamation + vbOKOnly
End Sub

Private Sub R_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, 0, 0
End Sub

Private Sub L_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, 0, 0
End Sub

Private Sub Bt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, 0, 0
End Sub

Private Sub Timer1_Timer()
If RG <= 255 - 16 Then
    RG = RG + 16
    R.BackColor = RGB(RG, RG, RG)
Else
    R.BackColor = vbWhite
    Timer1 = False
End If
End Sub

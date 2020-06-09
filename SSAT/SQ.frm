VERSION 5.00
Begin VB.Form SQ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "神器"
   ClientHeight    =   4848
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9924
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   24
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SQ.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4848
   ScaleWidth      =   9924
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton SM 
      Caption         =   "说明"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   528
      Left            =   8640
      MouseIcon       =   "SQ.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   240
      Width           =   1212
   End
   Begin VB.CommandButton Back 
      Caption         =   "<<返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   528
      Left            =   120
      MouseIcon       =   "SQ.frx":1114
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   240
      Width           =   1812
   End
   Begin VB.Frame F 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5266
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10000
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label"
         Height          =   1092
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   3240
         Width           =   3012
      End
      Begin VB.Label WH 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   72
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1464
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   1440
         Width           =   1464
      End
      Begin VB.Label BG 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   3852
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3012
      End
   End
End
Attribute VB_Name = "SQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OldCX As Long
Dim TotaL As Long

Const LbWid As Long = 2666
Const Jx As Long = 666

Private Sub Back_Click()
Unload Me
End Sub

Private Sub Back_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Back_Click
End If
End Sub

Private Sub Back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Back
    Form_MouseMove Button, Shift, X + .Left, Y + .Top
End With
End Sub

Private Sub SM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With SM
    Form_MouseMove Button, Shift, X + .Left, Y + .Top
End With
End Sub

Private Sub F_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With F
    Form_MouseMove Button, Shift, X + .Left, Y + .Top
End With
End Sub

Private Sub Form_Load()
TotaL = 7 * Jx + 6 * LbWid - Me.Width + Jx * 4 / 3
F.Width = 5 * Jx + 6 * LbWid + Jx * 4 / 3
Dim k As Integer
Dim Cl(), CP(), CT()
Cl = Array(&HFF0000, &HEF00&, &HFF, &H900090, &H50C000, &H0&)
CP = Array("完成游戏" & Chr(10) & "基本关卡", "右击主菜单" & Chr(10) & "标题6次", "让游戏溢出", "第十三关" & Chr(10) & "向左走", "使用秘籍", "集齐全部" & Chr(10) & "六个神器")
CT = Array("∑", "∷", "≡", "⊙", "∮", "?")
For k = 0 To 5
    If k <> 0 Then
        Load Label(k)
        Load WH(k)
        Load BG(k)
    End If
    With Label(k)
        .Caption = CP(k)
        .Top = Label(0).Top
        .Width = LbWid
        .Left = k * Jx + k * LbWid + Jx * 2 / 3
        .Visible = True
    End With
    With WH(k)
        .BackColor = Cl(k)
        .Top = WH(0).Top
        .Left = Label(k).Left + (Label(k).Width - .Width) / 2
        .Caption = "?"
        .Visible = True
    End With
    With BG(k)
        If GetSetting("iGlope", "woqu", "sq" & (k + 1), 0) = 1 Then
            .BackColor = &HCC0000
            .Caption = "已解锁"
            .ForeColor = vbWhite
            WH(k).Caption = CT(k)
        Else
            .BackColor = RGB(230, 230, 230)
            .Caption = "未解锁"
            .ForeColor = 0
        End If
        .Top = BG(0).Top
        .Left = Label(k).Left - Jx / 3
        .Width = LbWid + Jx * 2 / 3
        .Visible = True
    End With
Next k
End Sub

Private Sub RCX(cx As Long)
If cx <> OldCX Then
    OldCX = cx
    F.Left = Jx - cx / Me.Width * TotaL
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RCX Int(X)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell App.Path & "\" & App.EXEName, vbNormalFocus
End
End Sub

Private Sub SM_Click()
MsgBox "每个神器都有一个解锁条件，达成条件之后你将获得该神器。神器是收集品，没有功能。", vbInformation + vbOKOnly, "说明"
End Sub

Private Sub SM_KeyPress(KeyAscii As Integer)
Back_KeyPress KeyAscii
End Sub

Private Sub WH_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
With WH(Index)
    F_MouseMove Button, Shift, X + .Left, Y + .Top
End With
End Sub

Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
With Label(Index)
    F_MouseMove Button, Shift, X + .Left, Y + .Top
End With
End Sub

Private Sub BG_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
With BG(Index)
    F_MouseMove Button, Shift, X + .Left, Y + .Top
End With
End Sub

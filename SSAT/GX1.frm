VERSION 5.00
Begin VB.Form GX1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "更新"
   ClientHeight    =   2928
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   6036
   Icon            =   "GX1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2928
   ScaleWidth      =   6036
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Bt 
      Caption         =   "是"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   1
      Left            =   4560
      TabIndex        =   1
      Top             =   2280
      Width           =   1332
   End
   Begin VB.CommandButton Bt 
      Caption         =   "是"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   2280
      Width           =   1332
   End
   Begin VB.Label CP 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1368
   End
End
Attribute VB_Name = "GX1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim Mode As Boolean, nM As Boolean

Private Sub Bt_Click(Index As Integer)
If Mode = False Then
    Bt(0).Visible = False
    Bt(1).Visible = False
    Me.MousePointer = 11
    CP = "正在从 V1.5 更新到 V1.6 ..."
    DoEvents
    Sleep 6000
    CP = "更新失败！请检查网络连接。" & Chr(10) & "即将进入 我去冒险V1.7 ..."
    DoEvents
    Sleep 6000
    CP = "V1.7 新推出了“商店”功能。现在就去看看吧？"
    Bt(0).Visible = True
    Bt(1).Visible = True
    Bt(0).Caption = "否"
    Bt(1).Caption = "否"
    Mode = True
    Me.MousePointer = 0
Else
    Me.MousePointer = 11
    Bt(0).Visible = False
    Bt(1).Visible = False
    CP = "正在进入商店..."
    DoEvents
    Sleep 6000
    Shop.Show
    nM = True
    Unload Me
End If
End Sub

Private Sub Form_Load()
CP.Width = Me.Width
CP.Height = Bt(0).Top
CP = "检测到《我去冒险》有重要更新。是否现在下载V1.6？"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not nM Then
    Cancel = 1
    MsgBox "更新过程中无法关闭窗口。请完成更新。", vbExclamation + vbOKOnly, "There's no Easter Eggs here. Turn away. "
End If
End Sub

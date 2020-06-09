VERSION 5.00
Begin VB.Form LvFPS 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "第一人称FPS"
   ClientHeight    =   6528
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10416
   ForeColor       =   &H00000000&
   Icon            =   "LvFPS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   6528
   ScaleWidth      =   10416
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   " 敌人"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   1092
      Left            =   4440
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   2400
      Width           =   1812
   End
End
Attribute VB_Name = "LvFPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Music.ModeSwitch "fps"
Show
Me.FontSize = 20
Print "敌人出现了！赶快单击鼠标左键开火！"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
End
End Sub

Private Sub Label1_Click()
SaveSetting "iGlope", "woqu", "j4", 1
MsgBox "你通过了奖励关卡“第一人称FPS”。", vbOKOnly + vbInformation, "恭喜！"
Unload Me
End Sub

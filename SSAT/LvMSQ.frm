VERSION 5.00
Begin VB.Form LvMSQ 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "墨水池"
   ClientHeight    =   6528
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10416
   Icon            =   "LvMSQ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6528
   ScaleWidth      =   10416
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label MSQ 
      BackColor       =   &H00000000&
      Height          =   1692
      Left            =   3960
      TabIndex        =   0
      ToolTipText     =   "没错，这个墨水池什么用也没有！"
      Top             =   2280
      Width           =   2412
   End
End
Attribute VB_Name = "LvMSQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ND As Byte
Dim fX As Long, fY As Long

Private Sub Form_Load()
Me.ForeColor = vbWhite
Me.DrawWidth = 6
ND = 255
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ND < 255 Then
    ND = ND + 5
End If
Me.ForeColor = RGB(ND, ND, ND)
Line (X, Y)-(fX, fY)
fX = X
fY = Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "iGlope", "woqu", "j5", 1
MsgBox "你通过了奖励关卡“墨水池”。", vbOKOnly + vbInformation, "没错，离开就是答案！"
Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
End
End Sub

Private Sub MSQ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ND = 0
fX = X + MSQ.Left
fY = Y + MSQ.Top
End Sub


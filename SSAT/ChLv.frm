VERSION 5.00
Begin VB.Form ChLv 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "我去冒险"
   ClientHeight    =   4356
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   22.8
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ChLv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4356
   ScaleWidth      =   6480
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Focuser 
      Height          =   576
      Left            =   4680
      TabIndex        =   1
      Top             =   -666
      Width           =   972
   End
   Begin VB.Timer Dr 
      Interval        =   666
      Left            =   4080
      Top             =   3360
   End
   Begin VB.Timer TP 
      Enabled         =   0   'False
      Interval        =   166
      Left            =   4920
      Top             =   2880
   End
   Begin VB.TextBox Pad 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   2214
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   2052
   End
End
Attribute VB_Name = "ChLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nM As Boolean
Dim TopLv As Integer, Lv As Integer
Dim HC As Byte

Private Sub Dr_Timer()
Dr.Tag = 1 - Val(Dr.Tag)
PaintPicture LoadPicture(App.Path & "\fxj" & Dr.Tag & ".jpg"), 0, 3000, 2000, 1454
End Sub

Private Sub Form_Load()
Print "请选择关卡："
Dr_Timer
TopLv = GetSetting("iGlope", "woqu", "lv", 1)
Lv = TopLv
Renew
End Sub

Private Sub Renew()
Select Case Lv
    Case 1 To 3
        Pad = Lv
    Case 4
        Pad = "3.5"
    Case 5 To 9
        Pad = Lv - 1
    Case 10
        Pad = "8.5"
    Case 11 To 17
        Pad = Lv - 2
    Case 18
        Pad = "Boss"
End Select
If Lv <= TopLv Then
    Pad.Enabled = True
    Pad.BackColor = vbWhite
Else
    Pad.Enabled = False
    Pad.BackColor = RGB(236, 236, 236)
End If
End Sub

Private Sub Focuser_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27 'ESC
        Unload Me
        Exit Sub
    Case 38 'W
        With TP
            .Tag = 1
            .Enabled = True
        End With
        HC = 0
    Case 40 'S
        With TP
            .Tag = -1
            .Enabled = True
        End With
        HC = 0
    Case 13 'Enter
        If Lv <= TopLv Then
            Select Case Lv
                Case 1
                    Lv1.Show
                Case 2
                    Lv2.Show
                Case 3
                    Lv3.Show
                Case 4
                    Lv3_5.Show
                Case 5
                    Lv4.Show
                Case 6
                    Lv5.Show
                Case 7
                    Lv6.Show
                Case 8
                    Lv7.Show
                Case 9
                    Lv8.Show
                Case 10
                    Lv8_5.Show
                Case 11
                    Lv9.Show
                Case 12
                    Lv10.Show
                Case 13
                    Lv11.Show
                Case 14
                    Lv12.Show
                Case 15
                    Lv13.Show
                Case 16
                    Lv14.Show
                Case 17
                    Lv15.Show
                Case 18
                    LvBoss.Show
            End Select
            Music.ModeSwitch "lv"
            nM = True
            Unload Me
            Exit Sub
        Else
            MsgBox "这一关尚未解锁。请先完成第" & TopLv & "关。", vbExclamation + vbOKOnly
        End If
    Case Else
        Exit Sub
End Select
Lv = QJ(Lv + TP.Tag)
Renew
End Sub

Private Sub Focuser_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 38 'W
        TP.Enabled = False
    Case 40 'S
        TP.Enabled = False
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not nM Then
    Shell App.Path & "\" & App.EXEName, vbNormalFocus
    End
End If
End Sub

Private Sub Pad_GotFocus()
Focuser.SetFocus
End Sub

Private Sub TP_Timer()
If HC <= 6 Then
    HC = HC + 1
Else
    TopLv = QJ(TopLv + TP.Tag)
    Renew
End If
End Sub

Private Function QJ(Number As Integer) As Integer
QJ = Number
If QJ > 18 Then QJ = 18
If QJ < 1 Then QJ = 1
End Function

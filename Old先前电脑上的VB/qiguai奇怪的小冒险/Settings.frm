VERSION 5.00
Begin VB.Form Settings 
   Caption         =   "设置"
   ClientHeight    =   3096
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5316
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3096
   ScaleWidth      =   5316
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Flash 
      Left            =   3360
      Top             =   600
   End
   Begin VB.Timer Ji 
      Interval        =   1000
      Left            =   4800
      Top             =   1680
   End
   Begin VB.CommandButton Customize 
      Caption         =   "自定义背景音乐"
      Height          =   1092
      Left            =   5760
      TabIndex        =   3
      ToolTipText     =   "被你发现了！"
      Top             =   720
      Width           =   2052
   End
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   1680
      Top             =   1440
   End
   Begin VB.VScrollBar Tiao 
      Height          =   2292
      Left            =   3960
      Max             =   100
      TabIndex        =   2
      Top             =   240
      Width           =   492
   End
   Begin VB.CheckBox Check1 
      Caption         =   "程序猿模式"
      Enabled         =   0   'False
      Height          =   492
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   1452
   End
   Begin VB.CheckBox Music 
      Caption         =   "音乐"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   876
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Value           =   1  'Checked
      Width           =   2292
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   5040
      X2              =   4680
      Y1              =   120
      Y2              =   2760
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   2520
      X2              =   2640
      Y1              =   2880
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   120
      X2              =   2640
      Y1              =   1800
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   3360
      X2              =   2640
      Y1              =   120
      Y2              =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "右边这个滚动条控制音量大小。怎样，很别致吧？ "
      Height          =   1332
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   972
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S As Byte
Public MusicOff As Boolean
Dim a As Byte

Private Sub Customize_Click()
Theme.P.URL = InputBox("输入音乐的路径。记得带上尾缀命！", "自定义背景音乐", App.Path & "\R\1.wma")
End Sub

Private Sub Flash_Timer()
With Settings
    .Width = .Width + a
End With
a = a + 1
If a = 100 Then
    Flash.Interval = 0
    a = 0
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.Visible = True
End Sub

Private Sub Ji_Timer()
S = S + 1
Select Case S
    Case 10
        MsgBox "你不觉得这个“设置”窗口有些蹊跷吗？", vbOKOnly, ""
    Case 20
        MsgBox "站住别动！别关“设置”窗口！等20秒！", vbOKOnly, ""
    Case 40
        Flash.Interval = 20
        S = 0
        Ji.Interval = 0
End Select
End Sub

Private Sub Music_Click()
MusicOff = Music.Value - 1
End Sub

Private Sub Tiao_Change()
Theme.P.Settings.volume = 100 - Tiao.Value
End Sub

Private Sub Timer_Timer()
Theme.P.Settings.volume = 100 - Tiao.Value
End Sub

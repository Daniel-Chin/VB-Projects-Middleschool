VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "倒计时"
   ClientHeight    =   8676
   ClientLeft      =   192
   ClientTop       =   516
   ClientWidth     =   13968
   LinkTopic       =   "Form1"
   ScaleHeight     =   8676
   ScaleWidth      =   13968
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer3 
      Left            =   9120
      Top             =   3120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "更改闹铃音乐！"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "启用音乐"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   5880
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Left            =   10200
      Top             =   4920
   End
   Begin VB.Timer Timer1 
      Left            =   10200
      Top             =   4320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2760
      TabIndex        =   5
      Top             =   3360
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      TabIndex        =   2
      Text            =   "01"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      TabIndex        =   1
      Text            =   "00"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   0
      Text            =   "00"
      Top             =   2280
      Width           =   1095
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   1215
      Left            =   3120
      TabIndex        =   6
      Top             =   6840
      Width           =   1575
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2773
      _cy             =   2138
   End
   Begin VB.Label Label2 
      Caption         =   "："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      TabIndex        =   3
      Top             =   2280
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time As Long, time2 As Long, a As Integer

Private Sub Check1_Click()
If Check1.Value = 0 Then
    WindowsMediaPlayer1.URL = App.Path & "\null.wmv"
End If
End Sub

Private Sub Command1_Click()
If Timer1.Interval = 0 And Timer2.Interval = 0 Then
    Timer1.Interval = 1000
    Command1.Caption = "停止"
Else
    Command1.Caption = "开始"
    Timer1.Interval = 0
    Timer2.Interval = 0
    WindowsMediaPlayer1.Controls.stop
    Timer3.Interval = 0
    Form1.BackColor = RGB(205, 205, 205)
End If
Cls
End Sub

Private Sub Command2_Click()
Check1.Value = 1
WindowsMediaPlayer1.URL = InputBox("请输入地址", "自定义音乐", App.Path & "\1.wmv")
End Sub

Private Sub Form_Load()
a = 1
WindowsMediaPlayer1.Top = 0 - WindowsMediaPlayer1.Height - 20
WindowsMediaPlayer1.Left = 0 - WindowsMediaPlayer1.Width - 20
End Sub

Private Sub Timer1_Timer()
Text3 = Val(Text3) - 1
If Text3 = 0 Then Text3 = "00"
If Text3 = -1 Then Text2 = Val(Text2) - 1: Text3 = 59
If Text2 = -1 Then Text1 = Val(Text1) - 1: Text2 = 59
If Text1 = -1 Then Text1 = "00": Text2 = "00": Text3 = "00": Timer2.Interval = 500: Timer1.Interval = 0: WindowsMediaPlayer1.Controls.currentPosition = 0: WindowsMediaPlayer1.Controls.play: Timer3.Interval = 500
Form1.Caption = Text1 & ":" & Text2 & ":" & Text3 & "倒计时"
End Sub

Private Sub Timer2_Timer()
time2 = time2 + 1
Beep
If time2 > 50 Then Timer2.Interval = 0: time2 = 0
End Sub
Private Sub Timer3_Timer()
a = -a
If a = 1 Then
    Form1.BackColor = RGB(255, 0, 0)
Else
    Form1.BackColor = RGB(205, 205, 205)
End If
End Sub

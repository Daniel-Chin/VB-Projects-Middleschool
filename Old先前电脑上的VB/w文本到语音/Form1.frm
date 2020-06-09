VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4224
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11052
   LinkTopic       =   "Form1"
   ScaleHeight     =   4224
   ScaleWidth      =   11052
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3480
      Top             =   3000
   End
   Begin WMPLibCtl.WindowsMediaPlayer p 
      Height          =   1452
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   3012
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
      _cx             =   5313
      _cy             =   2561
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(), c(), w As Byte, tms As Integer
Const tp As Byte = 4

Private Sub Form_Load()
tms = 0
Me.FontSize = 18
a = Array(12.1, 5, 6, 2, 0, 13.7, 10.9, 18.8, 8, 1.1, 7, 0, 6, 9, 16, 21.2, 15, 4, 10.1, 14, 13, 26, 15, 20, 0, 18, 3, 18.9, 7.9, 16.9, 30.9, 28, 25, 24, 6, 14.9, 29.9, 22, 0, 18.9, 8, 23, 26, 31.9, 27, 33, 29, 34, 37, 0, 36, 6, 34.9, 39, 38, 40)
c = Array("朱", "易", "您", "好", "，", "我", "是", "宙", "斯", "一", "代", "。", "你", "能", "清", "楚", "地", "听", "见", "我", "在", "讲", "的", "话", "，", "证", "明", "宙", "斯", "已", "经", "学", "会", "了", "您", "的", "嗓", "音", "。", "宙", "斯", "即", "将", "被", "秦", "楠", "枫", "完", "成", "，", "请", "您", "拭", "目", "以", "待", "。")
For i = 1 To tp
    Load p(i)
Next i
For i = 0 To tp
    With p(i)
        .URL = App.Path & "\zhoce.wma"
        .settings.volume = 0
        .Top = -5000
        .Visible = False
    End With
Next i
Timer1.Interval = 400
End Sub

Private Sub Timer1_Timer()
w = (w + 1) Mod (tp + 1)
Dim b As Double
On Error Resume Next
    b = a(tms) * 2 + 3
    p(w).Controls.currentPosition = b
p(w).settings.volume = 100
tms = tms + 1
If tms = 60 Then
    Timer1.Interval = 0
    For i = 0 To tp
        p(i).Controls.stop
    Next i
End If
Print c(tms - 1);
If tms Mod 20 = 0 Then
    Print ""
End If
End Sub

Function DD(bb As String) As String
For i = Len(bb) To 1 Step -1
    DD = DD & Mid(bb, i, 1)
Next i
End Function

Sub fuck()
For i = 0 To tp - 1
    p(i).Controls.stop
Next i
p(tp).Controls.currentPosition = 0
End Sub


VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000012&
   ClientHeight    =   6348
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11328
   LinkTopic       =   "Form5"
   ScaleHeight     =   6348
   ScaleWidth      =   11328
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   5160
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6000
      Top             =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   25.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   2892
      Left            =   2040
      TabIndex        =   0
      Top             =   1920
      Width           =   8172
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time As Long
Dim nmal As Boolean
Private Sub Form_Unload(Cancel As Integer)
If nmal = True Then
Else
    End
End If
End Sub
Private Sub Form_Load()
Form11.WindowsMediaPlayer1.URL = App.Path & "\.W.o.w_files\ahahah.mp3"
time = -10
nmal = False
Form11.WindowsMediaPlayer1.settings.volume = 0
End Sub
Private Sub Timer1_Timer()
Label1.Left = (Form5.Width - Label1.Width) / 2
Label1.Top = (Form5.Height - Label1.Height) / 2
time = time + 1
Select Case time
    Case -6
        Label1.ForeColor = RGB(50, 50, 50): Label1 = "联邦军被击败。"
    Case -5
        Label1.ForeColor = RGB(100, 100, 100)
    Case -4
        Label1.ForeColor = RGB(150, 150, 150)
    Case -3
        Label1.ForeColor = RGB(200, 200, 200)
    Case -2
        Label1.ForeColor = RGB(255, 255, 255)
    Case 0
        Label1 = "我们败局已定。"
    Case 4
        Label1 = "我当时只是个新兵。我加入了“希望”计划。"
    Case 10
        Label1 = "这是最后一搏。"
    Case 12
        Label1 = "我和三十几个战友被改造成超级战士。我们有钢筋加固的骨头，能一拳穿透坦克装甲的拳头。"
    Case 16
        Label1 = "比猎豹更快，比狮子更凶猛，我们不再是人。我们已只是武器。"
    Case 20
        Label1 = "那波星，我们的第一个战场，我们被命令去摧毁帝国基地。"
    Case 23
        Label1 = "这是场屠杀。而被屠杀者是我们。我们面对的是四十万名帝国武士，我们一个又一个地倒下了。"
    Case 28
        Label1 = "而我，幸存了。"
    Case 31
        Label1 = "他们知道我没死。"
    Case 33
        Label1 = "我成了猎物。"
    Case 35
        Label1 = "他们是猎人。"
    Case 37
        Label1 = "新的旅途展开了。"
    Case 40
        Label1 = "新的旅途展开了。。": Label1.ForeColor = RGB(200, 200, 200)
    Case 41
        Label1 = "新的旅途展开了。。。": Label1.ForeColor = RGB(150, 150, 150)
    Case 42
        Label1 = "新的旅途展开了。。。。": Label1.ForeColor = RGB(100, 100, 100)
    Case 43
        Label1 = "新的旅途展开了。。。。": Label1.ForeColor = RGB(50, 50, 50)
    Case 44
        Label1 = "新的旅途展开了。。。。": Label1.ForeColor = RGB(0, 0, 0)
    Case 47
        nmal = True
        Unload Form5
        Form6.Visible = True
End Select
End Sub

Private Sub Timer2_Timer()
Form11.WindowsMediaPlayer1.settings.volume = Form11.WindowsMediaPlayer1.settings.volume + 1
If Form11.WindowsMediaPlayer1.settings.volume = 100 Then Timer2.Interval = 0
End Sub

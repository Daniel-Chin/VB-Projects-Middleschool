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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   5160
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   6000
      Top             =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "����"
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
        Label1.ForeColor = RGB(50, 50, 50): Label1 = "����������ܡ�"
    Case -5
        Label1.ForeColor = RGB(100, 100, 100)
    Case -4
        Label1.ForeColor = RGB(150, 150, 150)
    Case -3
        Label1.ForeColor = RGB(200, 200, 200)
    Case -2
        Label1.ForeColor = RGB(255, 255, 255)
    Case 0
        Label1 = "���ǰܾ��Ѷ���"
    Case 4
        Label1 = "�ҵ�ʱֻ�Ǹ��±����Ҽ����ˡ�ϣ�����ƻ���"
    Case 10
        Label1 = "�������һ����"
    Case 12
        Label1 = "�Һ���ʮ����ս�ѱ�����ɳ���սʿ�������иֽ�ӹ̵Ĺ�ͷ����һȭ��͸̹��װ�׵�ȭͷ��"
    Case 16
        Label1 = "���Ա����죬��ʨ�Ӹ����ͣ����ǲ������ˡ�������ֻ��������"
    Case 20
        Label1 = "�ǲ��ǣ����ǵĵ�һ��ս�������Ǳ�����ȥ�ݻٵ۹����ء�"
    Case 23
        Label1 = "���ǳ���ɱ��������ɱ�������ǡ�������Ե�����ʮ�����۹���ʿ������һ����һ���ص����ˡ�"
    Case 28
        Label1 = "���ң��Ҵ��ˡ�"
    Case 31
        Label1 = "����֪����û����"
    Case 33
        Label1 = "�ҳ������"
    Case 35
        Label1 = "���������ˡ�"
    Case 37
        Label1 = "�µ���;չ���ˡ�"
    Case 40
        Label1 = "�µ���;չ���ˡ���": Label1.ForeColor = RGB(200, 200, 200)
    Case 41
        Label1 = "�µ���;չ���ˡ�����": Label1.ForeColor = RGB(150, 150, 150)
    Case 42
        Label1 = "�µ���;չ���ˡ�������": Label1.ForeColor = RGB(100, 100, 100)
    Case 43
        Label1 = "�µ���;չ���ˡ�������": Label1.ForeColor = RGB(50, 50, 50)
    Case 44
        Label1 = "�µ���;չ���ˡ�������": Label1.ForeColor = RGB(0, 0, 0)
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

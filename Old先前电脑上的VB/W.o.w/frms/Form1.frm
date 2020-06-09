VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "World of War 客户端V1.4.7"
   ClientHeight    =   3804
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6024
   LinkTopic       =   "Form1"
   ScaleHeight     =   3804
   ScaleWidth      =   6024
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   5160
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   5160
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5040
      Top             =   1680
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FF00&
      Height          =   372
      Left            =   -50
      Top             =   3480
      Width           =   132
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   492
      Left            =   3600
      TabIndex        =   1
      Top             =   2880
      Width           =   2172
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "正在初始化环境"
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   1080
      TabIndex        =   0
      Top             =   2880
      Width           =   3972
   End
   Begin VB.Image Image1 
      Height          =   2964
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Long
Dim b As Long
Dim nmal As Boolean

'Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then Form7.Visible = True
'End Sub

Private Sub Form_Load()
'If App.Path <> "C:\W.o.w" Then
'    MsgBox "对不起，请将W.o.w文件夹放在C盘根目录下", vbOKOnly, "运行失败！"
'    End
'End If
nmal = False
Form11.Visible = True
a = 0
b = 0
'nmal = True
'Form9.Visible = True
'Unload Form1
End Sub
Private Sub 你妈()
Form4.Visible = True
nmal = True
Unload Form1
End Sub
Public Sub 八路()
Form8.Visible = True
nmal = True
Unload Form1
End Sub
Public Sub 六六大顺()
Form6.Visible = True
nmal = True
Unload Form1
End Sub
Private Sub 卡卡()
Form10.Visible = True
nmal = True
Unload Form1
End Sub
Private Sub 哇操()
Form5.Visible = True
nmal = True
Unload Form1
End Sub
Private Sub 蛋蛋()
Form7.Visible = True
nmal = True
Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
If nmal = True Then
Else
    End
End If
End Sub

Private Sub Timer1_Timer()
a = a + 1
If a Mod 4 = 0 Then
    Label2 = ""
End If
If a Mod 4 = 1 Then
    Label2 = "."
End If
If a Mod 4 = 2 Then
    Label2 = ".."
End If
If a Mod 4 = 3 Then
    Label2 = "..."
End If
If Shape1.Width > 2999 And b < 12 Then
Else
    Shape1.Width = Shape1.Width + 100
End If
Form1.Width = 6240
Form1.Height = 4344
End Sub

Private Sub Timer2_Timer()
b = b + 1
If b = 3 Then Label1 = "正在加载地图"
If b = 8 Then Label1 = "正在加载道具"
If b = 13 Then Label1 = "正在检测您安装的flash版本"
If b = 14 Then Label1 = "正在启动游戏"
If b = 17 Then Form3.Visible = True: nmal = True: Unload Form1
End Sub
Private Sub Timer3_Timer()
Shape1.Width = Shape1.Width + 8
If Shape1.Width > 3000 And b < 12 Then Shape1.Width = 3000
End Sub

VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "挑战模式"
   ClientHeight    =   7536
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8652
   LinkTopic       =   "Form1"
   ScaleHeight     =   7536
   ScaleWidth      =   8652
   Begin VB.Timer Timer4 
      Left            =   6600
      Top             =   2160
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0000FF00&
      Caption         =   "Command9"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   0
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   4092
   End
   Begin VB.Timer Timer3 
      Left            =   6960
      Top             =   3960
   End
   Begin VB.Timer Timer2 
      Left            =   6720
      Top             =   3240
   End
   Begin VB.TextBox Text15 
      Height          =   492
      Left            =   9000
      TabIndex        =   20
      Text            =   "0"
      Top             =   720
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.CommandButton Command6 
      Caption         =   "设置和游戏模式"
      Height          =   492
      Left            =   6480
      TabIndex        =   19
      Top             =   120
      Width           =   1812
   End
   Begin VB.TextBox Text12 
      Height          =   492
      Left            =   9960
      TabIndex        =   18
      Text            =   "0"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.TextBox Text11 
      Height          =   372
      Left            =   9960
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text14 
      Height          =   264
      Left            =   10560
      TabIndex        =   14
      Text            =   "0"
      Top             =   3240
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Height          =   492
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2040
      Width           =   492
   End
   Begin VB.TextBox Text9 
      Height          =   372
      Left            =   12840
      TabIndex        =   12
      Text            =   "-25"
      Top             =   4320
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Height          =   372
      Left            =   12600
      TabIndex        =   11
      Text            =   "25"
      Top             =   3600
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text7 
      Height          =   372
      Left            =   12600
      TabIndex        =   10
      Text            =   "30"
      Top             =   3000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text6 
      Height          =   372
      Left            =   12600
      TabIndex        =   9
      Text            =   "28"
      Top             =   2520
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text5 
      Height          =   372
      Left            =   12960
      TabIndex        =   8
      Text            =   "20"
      Top             =   1800
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text4 
      Height          =   372
      Left            =   12840
      TabIndex        =   7
      Text            =   "-30"
      Top             =   1200
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF0000&
      Height          =   732
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   972
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Height          =   732
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   852
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Height          =   900
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Height          =   372
      Left            =   3360
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1692
   End
   Begin VB.TextBox Text3 
      Height          =   1212
      Left            =   9600
      TabIndex        =   2
      Text            =   "0"
      Top             =   2160
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   12840
      TabIndex        =   1
      Text            =   "-20"
      Top             =   600
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   13200
      TabIndex        =   0
      Text            =   "-25"
      Top             =   120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Timer Timer1 
      Left            =   11160
      Top             =   6240
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Warning"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1092
      Left            =   720
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   4092
   End
   Begin VB.Label Label3 
      Caption         =   "time"
      Height          =   612
      Left            =   9120
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "Check"
      Height          =   372
      Left            =   9720
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   564
      Left            =   0
      Top             =   5280
      Width           =   5892
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   5892
      Left            =   5760
      Top             =   0
      Width           =   492
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim achi_load As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim jiasudu As Double
Dim position As POINTAPI
Dim t_34 As Long
Dim yici As Boolean      '后面timer1是否打开timer2要用。
Dim yuanlai_boss_h As Long
Dim yuanlai_boss_w As Long
Dim boss_biandale As Boolean
Dim achi_cap As String
Private Sub Command5_Click()
If Text14 = 0 Then
    Timer1.Interval = 15
End If
jiasudu = 1.002
End Sub

Private Sub Command6_Click()
Form3.Visible = True
End Sub
Private Sub Command9_Click()
If Form1.achi_achibtm = False Then          '成就：按钮
    Form1.achi_achibtm = True
    achi_cap = "看到按钮就想点"
    Call achibb2
End If
End Sub

Private Sub Form_Load()
yuanlai_boss_h = Command2.Height
 yuanlai_boss_w = Command2.Width
Call 重置
End Sub
Private Sub achibb2()    'form2的成就下滑，bb是bubble的意思
Command9.Caption = "获得成就：" & achi_cap
achi_load = -40
Timer4.Interval = 30
End Sub

Private Sub Timer1_Timer()
Text3 = Text3 + 1
If Command1.Top < 0 Or Command1.Top > Shape2.Top - Command1.Height Then Text1 = -Text1      '这些控制方块反弹
If Command1.Left < 0 Or Command1.Left > Shape1.Left - Command1.Width Then Text2 = -Text2
If Command2.Top < 0 Or Command2.Top > Shape2.Top - Command2.Height Then Text4 = -Text4
If Command2.Left < 0 Or Command2.Left > Shape1.Left - Command2.Width Then Text5 = -Text5
If Command3.Top < 0 Or Command3.Top > Shape2.Top - Command3.Height Then Text6 = -Text6
If Command3.Left < 0 Or Command3.Left > Shape1.Left - Command3.Width Then Text7 = -Text7
If Command4.Top < 0 Or Command4.Top > Shape2.Top - Command4.Height Then Text8 = -Text8
If Command4.Left < 0 Or Command4.Left > Shape1.Left - Command4.Width Then Text9 = -Text9
Command1.Top = Command1.Top + Text1                              '方块移动
Command1.Left = Command1.Left + Text2
Command2.Top = Command2.Top + Text4
Command2.Left = Command2.Left + Text5
Command3.Top = Command3.Top + Text6
Command3.Left = Command3.Left + Text7
Command4.Top = Command4.Top + Text8
Command4.Left = Command4.Left + Text9
If Text3 < 250 Then                  '设定到什么时候不加速
    Text1 = Text1 * jiasudu                '加速
    Text2 = Text2 * jiasudu
    Text4 = Text4 * jiasudu
    Text5 = Text5 * jiasudu
    Text6 = Text6 * jiasudu
    Text7 = Text7 * jiasudu
    Text8 = Text8 * jiasudu
    Text9 = Text9 * jiasudu
End If
GetCursorPos position
ScreenToClient Me.hwnd, position
Command5.Top = position.y * Form1.x_s - 200
Command5.Left = position.x * Form1.x_s - 200
Call 碰到
If Text3 * 1 > Text10 * 1 Then Text10 = Text3
If Text14 = 1 Then
    Call MsgBox("你死了。", vbExclamation, "你死了！")
    Timer1.Interval = 0
    If boss_biandale = True Then
        Form1.achi_bosskengdie = True
        achi_cap = "Boss怎么这么肯爹？"
        Call achibb2
    End If
    Call 重置
End If
If Text3 > 3000 Then
    If yici = False Then
        Timer2.Interval = Timer1.Interval
        Timer3.Interval = 500
        Command3.Visible = False: Command3.Top = -1000: Command3.Left = -1000
        Command4.Visible = False: Command4.Top = -1000: Command4.Left = -1000
        Command1.Visible = False: Command1.Top = -1000: Command1.Left = -1000
        yici = True
    End If
End If
End Sub
Sub 碰到()
If Command5.Top > (Command1.Top - Command5.Height) And Command5.Top < (Command1.Top + Command1.Height) And Command5.Left < (Command1.Left + Command1.Width) And Command5.Left > (Command1.Left - Command5.Width) Then Text14 = 1
If Command5.Top > (Command2.Top - Command5.Height) And Command5.Top < (Command2.Top + Command2.Height) And Command5.Left < (Command2.Left + Command2.Width) And Command5.Left > (Command2.Left - Command5.Width) Then Text14 = 1
If Command5.Top > (Command3.Top - Command5.Height) And Command5.Top < (Command3.Top + Command3.Height) And Command5.Left < (Command3.Left + Command3.Width) And Command5.Left > (Command3.Left - Command5.Width) Then Text14 = 1
If Command5.Top > (Command4.Top - Command5.Height) And Command5.Top < (Command4.Top + Command4.Height) And Command5.Left < (Command4.Left + Command4.Width) And Command5.Left > (Command4.Left - Command5.Width) Then Text14 = 1
If Command5.Top > Shape2.Top - Command5.Height Or Command5.Left > Shape1.Left - Command5.Width Or Command5.Top < 0 Or Command5.Left < 0 Then Text14 = 1
End Sub
Sub 重置()
Timer2.Interval = 0
Timer3.Interval = 0
Label2.Visible = False
Command2.Width = yuanlai_boss_w
Command2.Height = yuanlai_boss_h
Text1 = -15
Text2 = -15
Text4 = -15
Text5 = 15
Text6 = 15
Text7 = 15
Text8 = 15
Text9 = -15
Text11 = 1
Text12 = 0
Command1.Top = 2900
Command1.Left = 2900
Command1.Visible = True
Command5.Top = 1800
Command5.Left = 1800
Command5.Visible = True
Command2.Top = 2900
Command2.Left = 120
Command2.Visible = True
Command3.Top = 120
Command3.Left = 120
Command3.Visible = True
Command4.Top = 0
Command4.Left = 2900
Command4.Visible = True
Text14 = 0
Text3 = 0
t_34 = 0
yici = False
boss_biandale = False
End Sub
Private Sub Timer2_Timer()
t_34 = t_34 + 1
Command2.Width = Command2.Width + 3
Command2.Height = Command2.Height + 3
If Command2.Left < 0 Then
    Command2.Left = Command2.Left + 10
End If
If Command2.Left + Command2.Width > Shape1.Left - 11 Then
    Command2.Left = Command2.Left - 10
End If
If Command2.Top < 0 Then
    Command2.Top = Command2.Top + 10
End If
If Command2.Top + Command2.Height > Shape2.Top - 11 Then
    Command2.Top = Command2.Top - 10
End If
If t_34 > 800 Then
    Timer2.Interval = 0
    Timer3.Interval = 0
    Label2.Visible = False
    Command2.Top = Command2.Top - 5
    boss_biandale = True
End If
End Sub

Private Sub Timer3_Timer()
If t_34 < 800 Then
    Label2.Visible = Not (Label2.Visible)
End If
End Sub

Private Sub Timer4_Timer()
Command9.Top = 40 - achi_load ^ 2
Command9.Visible = True
achi_load = achi_load + 1
If achi_load > 50 Then Timer2.Interval = 0: Command9.Visible = False
End Sub

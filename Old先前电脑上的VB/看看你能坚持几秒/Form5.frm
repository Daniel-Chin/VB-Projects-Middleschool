VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "无限模式"
   ClientHeight    =   7536
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8652
   LinkTopic       =   "Form1"
   ScaleHeight     =   7536
   ScaleWidth      =   8652
   Begin VB.TextBox Text15 
      Height          =   492
      Left            =   9000
      TabIndex        =   21
      Text            =   "0"
      Top             =   720
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.CommandButton Command6 
      Caption         =   "设置和游戏模式"
      Height          =   492
      Left            =   6600
      TabIndex        =   20
      Top             =   120
      Width           =   1812
   End
   Begin VB.TextBox Text12 
      Height          =   492
      Left            =   9960
      TabIndex        =   19
      Text            =   "0"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.TextBox Text11 
      Height          =   372
      Left            =   9960
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text10 
      Height          =   372
      Left            =   120
      TabIndex        =   17
      Text            =   "0"
      Top             =   6120
      Width           =   732
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   492
   End
   Begin VB.TextBox Text9 
      Height          =   372
      Left            =   12840
      TabIndex        =   12
      Text            =   "-28"
      Top             =   4320
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Height          =   372
      Left            =   12600
      TabIndex        =   11
      Text            =   "28"
      Top             =   3600
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text7 
      Height          =   372
      Left            =   12600
      TabIndex        =   10
      Text            =   "38"
      Top             =   3000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text6 
      Height          =   372
      Left            =   12600
      TabIndex        =   9
      Text            =   "30"
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
      Text            =   "-35"
      Top             =   1200
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF0000&
      Height          =   732
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   972
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Height          =   732
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   852
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Height          =   900
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Height          =   372
      Left            =   4200
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
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
      Text            =   "-10"
      Top             =   600
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   13200
      TabIndex        =   0
      Text            =   "-20"
      Top             =   120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Timer Timer1 
      Left            =   11160
      Top             =   6240
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
      Top             =   5520
      Width           =   6492
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   5892
      Left            =   6000
      Top             =   0
      Width           =   492
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim jiasudu As Double
Dim position As POINTAPI


Public Sub dimdimdim()
Dim achi_combo As Boolean
End Sub

Private Sub Command5_Click()
If Text14 = 0 Then
    Timer1.Interval = 15
End If
jiasudu = 0.1
End Sub
Private Sub Form_Load()
Call 重置
Call MsgBox("说明：这是无限模式。不能使用道具，双倍收获经验或金币。", vbExclamation, "说明")
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
    Text1 = Text1 + IIf(Text1 > 0, jiasudu, jiasudu * (-1))          '加速
    Text2 = Text2 + IIf(Text2 > 0, jiasudu, jiasudu * (-1))
    Text4 = Text4 + IIf(Text4 > 0, jiasudu, jiasudu * (-1))
    Text5 = Text5 + IIf(Text5 > 0, jiasudu, jiasudu * (-1))
    Text6 = Text6 + IIf(Text6 > 0, jiasudu, jiasudu * (-1))
    Text7 = Text7 + IIf(Text7 > 0, jiasudu, jiasudu * (-1))
    Text8 = Text8 + IIf(Text8 > 0, jiasudu, jiasudu * (-1))
    Text9 = Text9 + IIf(Text9 > 0, jiasudu, jiasudu * (-1))
GetCursorPos position
ScreenToClient Me.hwnd, position
Command5.Top = position.y * Form1.x_s - 200
Command5.Left = position.x * Form1.x_s - 200
Call 碰到
If Text3 * 1 > Text10 * 1 Then Text10 = Text3
If Text14 = 1 Then
    Label4 = "你死了，单击方块重新开始"
    Call MsgBox("你得了" & Text3 & "分。", vbExclamation, "你死了！")
    Call 重置
    Timer1.Interval = 0
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
Text1 = -20
Text2 = -10
Text4 = -15
Text5 = 20
Text6 = 15
Text7 = 10
Text8 = 18
Text9 = -18
Text11 = 1
Text12 = 0
Command1.Top = 4440
Command1.Left = 4200
Command1.Visible = True
Command5.Top = 2500
Command5.Left = 2500
Command5.Visible = True
Command2.Top = 3720
Command2.Left = 840
Command2.Visible = True
Command3.Top = 720
Command3.Left = 120
Command3.Visible = True
Command4.Top = 600
Command4.Left = 4560
Command4.Visible = True
Text14 = 0
Text3 = 0
End Sub

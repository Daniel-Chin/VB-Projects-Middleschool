VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "提醒：Minecraft学习"
   ClientHeight    =   2232
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   3672
   LinkTopic       =   "Form1"
   ScaleHeight     =   2232
   ScaleWidth      =   3672
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "返回设置去"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   3492
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   2280
      Top             =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "请勿再次提醒了"
      Height          =   252
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   2292
   End
   Begin VB.CommandButton Command2 
      Caption         =   "暂时不要"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "好的"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3492
   End
   Begin VB.Label Label1 
      Caption         =   "现在学习Minecraft知识么？"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3372
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public time As Long
Dim time1 As Long
Private Sub OpenUrl(tUrl As String)
    ShellExecute Me.hwnd, "Open", tUrl, 0, 0, 0
End Sub

Private Sub Command1_Click()
    OpenUrl "http://zh.minecraftwiki.net/wiki/Special:%E9%9A%8F%E6%9C%BA%E9%A1%B5%E9%9D%A2"
    Form1.Visible = False
End Sub

Private Sub Command2_Click()
Form1.Visible = False
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Form2.Visible = True
Form1.Visible = False
End Sub

Private Sub Timer1_Timer()
time1 = time1 + 1
If time1 > time Then
    time1 = 0
    Form1.Visible = True
    Beep
End If
End Sub

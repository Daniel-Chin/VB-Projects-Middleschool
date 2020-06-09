VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   4368
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8064
   LinkTopic       =   "Form1"
   ScaleHeight     =   4368
   ScaleWidth      =   8064
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   3252
   End
   Begin VB.Image Image5 
      Height          =   2724
      Left            =   3720
      Picture         =   "Form1.frx":0000
      Top             =   1200
      Width           =   4548
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "英雄联盟 快捷启动器"
      BeginProperty Font 
         Name            =   "华文行楷"
         Size            =   25.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   612
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   4932
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   480
      Picture         =   "Form1.frx":364A
      Top             =   3840
      Width           =   3240
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   480
      Picture         =   "Form1.frx":4697
      Top             =   3120
      Width           =   3240
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   480
      Picture         =   "Form1.frx":6265
      Top             =   2400
      Width           =   3240
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "Form1.frx":8943
      Top             =   1680
      Width           =   3240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub OpenUrl(tUrl As String)
    ShellExecute Me.hwnd, "Open", tUrl, 0, 0, 0
End Sub

Private Sub Form_Load()
Dim a As String, b As String
Open App.Path & "\1.txt" For Input As #1
    Input #1, a
    
    
Close
If a = "" Then
    b = InputBox("请输入您的昵称", "起一个好听点的名字吧", "我叫qnh")
    Open App.Path & "\1.txt" For Output As #2
        Print #2, b
    Close
    a = b
End If
Label2.Caption = "您好，玩家  " & a
End Sub

Private Sub Image1_Click()
    OpenUrl "G:\其他\英雄联盟\TCLS\Client.exe"
End Sub

Private Sub Image2_Click()
    OpenUrl "C:\Program Files\360Safebox\360safebox.exe"
End Sub

Private Sub Image3_Click()
    OpenUrl "http://www.lol.qq.com/"
End Sub

Private Sub Image4_Click()
    End
End Sub

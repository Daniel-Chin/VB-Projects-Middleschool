VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "开服器"
   ClientHeight    =   5568
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9636
   LinkTopic       =   "Form1"
   ScaleHeight     =   5568
   ScaleWidth      =   9636
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton H 
      BackColor       =   &H0080C0FF&
      Caption         =   "秦楠海的Server"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4332
      Left            =   5592
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   618
      Width           =   2892
   End
   Begin VB.CommandButton F 
      BackColor       =   &H00FF8080&
      Caption         =   "秦楠枫的Server"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4332
      Left            =   1152
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   618
      Width           =   2892
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub F_Click()
On Error Resume Next
Kill "F:\MC\I\1.6.4\server.properties"
FileCopy App.Path & "\qnf", "F:\MC\I\1.6.4\server.properties"
Shell "F:\MC\I\1.6.4\开始游戏.exe", vbNormalFocus
Shell "F:\MC\I\1.6.4\Server.exe", vbNormalFocus
End
End Sub

Private Sub H_Click()
On Error Resume Next
Kill "F:\MC\I\1.6.4\server.properties"
FileCopy App.Path & "\qnh", "F:\MC\I\1.6.4\server.properties"
Shell "F:\MC\I\1.6.4\Server.exe", vbNormalFocus
End
End Sub

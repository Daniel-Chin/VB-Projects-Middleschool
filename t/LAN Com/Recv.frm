VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1740
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5568
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   5568
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   1332
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   2172
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   1200
      TabIndex        =   3
      Text            =   "1875"
      Top             =   720
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      Height          =   372
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   1200
      TabIndex        =   1
      Text            =   "116.232.228.152"
      Top             =   240
      Width           =   1812
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   372
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   972
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1680
      Top             =   1080
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "ip"
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    '发送
    Dim ass As String
    ass = InputBox("要发送什么消息？")
    'ass = 49
    Winsock1.SendData ass
    Text3 = Text3 + vbCrLf + ass
End Sub

Private Sub Command2_Click()
    With Winsock1
        ' 采用UDP协议
        .Protocol = sckUDPProtocol
        '.RemoteHost = "192.168.0.151"
        '接收端程序所使用的端口
        .LocalPort = Val(Text2)
        '打开Socket
        .Bind
    End With
    Command2.Visible = False
    Command1.Enabled = True
    Text3 = Text3 + "Connection established"
End Sub

Private Sub Form_Load()
    Me.Caption = "Daniel1875服务端"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '窗口关闭时，先关闭Winsock1
    Winsock1.Close
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    '接收数据
    Dim strData As String
    Winsock1.GetData strData, vbString
    '将接收的显示在textBox1
    Text3 = Text3 + vbCrLf + strData
End Sub

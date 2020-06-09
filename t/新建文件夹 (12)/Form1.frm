VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4488
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   8628
   LinkTopic       =   "Form1"
   ScaleHeight     =   4488
   ScaleWidth      =   8628
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   3240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   4092
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4200
      Top             =   2040
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
      LocalPort       =   23333
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    '初始化Socket
    With Winsock1
        '采用UDP协议
        .Protocol = sckUDPProtocol
        '本机端口
        .LocalPort = 1875
        '打开Socket
        .Bind
    End With
    Text1.Text = ""
    Me.Caption = "UDP接收端程序"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '关闭窗口时先关闭socket
    Winsock1.Close
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    '接收数据
    Dim strData As String
    Winsock1.GetData strData, vbString
    '将接收的显示在textBox1
    Text1.Text = strData
End Sub

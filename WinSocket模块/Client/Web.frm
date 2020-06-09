VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Web 
   Caption         =   "Web"
   ClientHeight    =   4164
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   42
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4164
   ScaleWidth      =   7500
   StartUpPosition =   1  '所有者中心
   Begin MSWinsockLib.Winsock Sock 
      Left            =   6240
      Top             =   1920
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.TextBox Prompt 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3972
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6972
   End
End
Attribute VB_Name = "Web"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'以下是封装代码，请勿随意更改
Option Explicit

Dim IP As String
Dim Port As Integer

Public Sub Connect(Optional newIP As String = "", Optional newPort As Integer = -1)
If newPort <> -1 Then
    Port = newPort
End If
If newIP <> "" Then
    IP = newIP
End If
With Sock
    .Protocol = sckTCPProtocol
    .RemoteHost = IP
    .RemotePort = Port
    .Connect
End With
End Sub

Private Sub Form_Resize()
Prompt.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
Sock.Close
End
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Sock.GetData strData, vbString
    Pop "Data from Server: " & strData
    Dim FirstLen As Integer
    strData = Right(strData, Len(strData) - 1)
    FirstLen = InStr(strData, Chr(2))
    Do Until FirstLen = 0
        ReadData Left(strData, FirstLen - 1)
        strData = Right(strData, Len(strData) - InStr(strData, Chr(2)))
        FirstLen = InStr(strData, Chr(2))
    Loop
    ReadData strData
End Sub

Public Sub Pop(MSG As String)
Prompt = Prompt & MSG & vbCrLf
Prompt.SelStart = Len(Prompt)
Debug.Print MSG
End Sub

'以上是封装代码，请勿随意更改。
'以下是本模块的使用

Private Sub ReadData(strData As String)
'这个事件接收来自服务器的消息，收到的消息存在strData里
End Sub

Public Sub Send(strData As String)
'调用Send方法，向服务端发送文本消息。
On Error GoTo WebErr
DoEvents
Sock.SendData Chr(2) & strData
DoEvents
Pop "Data sent: " & strData
Exit Sub
WebErr:
Pop "哎呀，网络连接出现了问题！"
End Sub

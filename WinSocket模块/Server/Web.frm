VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Web 
   Caption         =   "WEB"
   ClientHeight    =   5412
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9852
   LinkTopic       =   "Form1"
   ScaleHeight     =   5412
   ScaleWidth      =   9852
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Prompt 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4572
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   4092
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   9480
      Top             =   360
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
      LocalPort       =   1875
   End
End
Attribute VB_Name = "Web"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'以下是封装代码，请勿随意更改
Option Explicit

Const MaxConnection As Integer = 999
Public Port As Integer
Private Gate(1 To MaxConnection) As Boolean

Public Sub Listen(Optional NewPort As Integer = -1)
If NewPort <> -1 Then
    Port = NewPort
End If
With Sock(0)
    .LocalPort = Port
    .Listen
End With
Pop "开始聆听。端口: " & Port
End Sub

Private Sub Form_Resize()
Prompt.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim k As Integer
For k = 1 To Sock.ubound
    If Gate(k) Then
        Sock(k).Close
    End If
Next k
Sock(0).Close
End Sub

Private Sub Sock_Close(Index As Integer)
Gate(Index) = False
Unload Sock(Index)
Pop String(Index - 1, "-") & "断开: ID " & Index
End Sub

Private Sub Sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim k As Integer
Do Until k = MaxConnection
    k = k + 1
    If Not Gate(k) Then
        Exit Do
    End If
Loop
If k = MaxConnection Then
    Pop "连接数量已满！有新的请求！"
Else
    Load Sock(k)
    Gate(k) = True
    Sock(k).Accept requestID
    Pop "连接: ID " & k
    DoEvents
End If
End Sub

Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Sock(Index).GetData strData, vbString
Pop String(Index - 1, "-") & "From ID " & Index & ": " + strData
Dim FirstLen As Integer
strData = Right(strData, Len(strData) - 1)
FirstLen = InStr(strData, Chr(2))
Do Until FirstLen = 0
    ReadData Index, Left(strData, FirstLen - 1)
    strData = Right(strData, Len(strData) - InStr(strData, Chr(2)))
    FirstLen = InStr(strData, Chr(2))
Loop
ReadData Index, strData
End Sub

Private Sub Pop(MSG As String)
Prompt = Prompt & MSG & vbCrLf
Prompt.SelStart = Len(Prompt)
Debug.Print MSG
End Sub

'以上是封装代码，请勿随意更改。
'以下是本模块的使用

Private Sub ReadData(Index As Integer, strData As String)
'这个事件接收来自客户端的消息，收到的消息存在strData里
'Index就是客户端的ID。
End Sub

Public Function IsGate(ID As Integer) As Boolean
'这个函数判断编号为ID的客户端是否连接着
IsGate = Gate(ID)
End Function

Public Sub Send(ToWho As Integer, strData As String)
'调用Send方法，向客户端发送文本消息。
'ToWho是目标客户端的编号，strData是消息内容
If Gate(ToWho) Then
    DoEvents
    Sock(ToWho).SendData Chr(2) & strData
    DoEvents
    Pop String(ToWho - 1, "-") & "To ID " & ToWho & ": " & strData
Else
    Pop "发送目标已断开连接: " & ToWho & "，无法发送" & strData
End If
End Sub

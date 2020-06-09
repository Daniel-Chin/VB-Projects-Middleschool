VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form form1 
   Caption         =   "联网模块"
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
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   6972
   End
   Begin MSWinsockLib.Winsock Sock 
      Left            =   6720
      Top             =   2160
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Private Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type

Dim IP As String
Dim Port As Integer

Dim IntPerS As Integer
Dim IntLastS As Integer
Dim LastS As Integer

Dim MyID As Integer

Dim Pinged As Boolean

Private Sub Form_Load()
Randomize
MsgBox "本游戏由iGlope开发。游戏机制不喻指任何国家的股票交易，仅供娱乐。" + vbCrLf + "这是一个多人联机游戏，服务器目前位于我家。游戏之前，请通知我打开服务器，向我索要服务器IP和端口号。"
IP = InputBox("服务器IP：", , GetSetting("iglope", "stocket", "ip", "192.168.0.151"))
Port = InputBox("端口号：", , GetSetting("iglope", "stocket", "port", "1875"))
SaveSetting "iglope", "stocket", "ip", IP
SaveSetting "iglope", "stocket", "port", Port
Move 0, 0
'IP = "192.168.0.151"
'Port = "1875"
With Sock
    .Protocol = sckTCPProtocol
    .RemoteHost = IP
    .RemotePort = Port
    .Connect
End With
Sleep 100
DoEvents
Send "ping", ""
Dim k As Integer
For k = 1 To 100
    Sleep 10
    DoEvents
    If Pinged Then
        Exit For
    End If
Next k
If Pinged Then
    Hall.Show
    Pinged = False
Else
    MsgBox "额，无法链接到服务器。", vbCritical
    Unload Me
    End
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Sock.Close
End
End Sub

Public Function NowMs() As Integer
Dim T As SYSTEMTIME
GetLocalTime T
NowMs = T.wMilliseconds
End Function

Public Sub Send(Operation As String, Detail)
DoEvents
On Error GoTo WebErr
Sock.SendData "`" & Second(Now) & " " & NowMs & " " & Operation & " " & Detail
DoEvents
Pop "Data sent: " & Operation & " " & Detail
Exit Sub
WebErr:
Pop "哎呀，网络连接出现了问题！"
End Sub

Public Sub Buy(BuyIndex As Integer)
Send "buy", BuyIndex
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Sock.GetData strData, vbString
    Pop "Data from Server: " + strData
    Dim FirstLen As Integer
    strData = Right(strData, Len(strData) - 1)
    FirstLen = InStr(strData, "`")
    Do Until FirstLen = 0
        ReadData Left(strData, FirstLen - 1)
        strData = Right(strData, Len(strData) - InStr(strData, "`"))
        FirstLen = InStr(strData, "`")
    Loop
    ReadData strData
End Sub

Private Sub ReadData(strData As String)
    Dim S() As String
    S = Split(strData, " ")
    Select Case S(2)
        Case "room"
            Hall.RenewRoom strData
        Case "play"
            Hall.Play
        Case "loadindex"
            GUI.LoadIndex Val(S(3))
        Case "gnip"
            Pinged = True
        Case "balance"
            GUI.RenewBalance S(3)
        Case "market"
            GUI.AddBuff Val(S(0)), Val(S(1)), strData
        Case "chat"
            GUI.Chat strData
        Case "gameover"
            GUI.GameOver
        Case "msgbox"
            MsgBox S(3), vbInformation, "来自服务器的消息"
        Case "end"
            MsgBox "服务器已经断开连接。"
            Unload Me
            End
        Case Else
            Stop
    End Select
End Sub

Public Sub Pop(MSG As String)
Prompt = Prompt & MSG & vbCrLf
'Prompt.BackColor = &HFF0000 - Prompt.BackColor
Prompt.SelStart = Len(Prompt)
Debug.Print MSG
End Sub

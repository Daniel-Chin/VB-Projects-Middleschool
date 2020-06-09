VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Web 
   Caption         =   "股市模拟服务端"
   ClientHeight    =   5412
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9852
   LinkTopic       =   "Form1"
   ScaleHeight     =   5412
   ScaleWidth      =   9852
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Disconnecter 
      Caption         =   "Disconnect"
      Height          =   372
      Left            =   1920
      TabIndex        =   5
      Top             =   3480
      Width           =   1452
   End
   Begin VB.TextBox Prompt 
      Height          =   4572
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   360
      Width           =   4092
   End
   Begin VB.CommandButton RenewRoomer 
      Caption         =   "RenewRoom"
      Height          =   372
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   1572
   End
   Begin VB.CommandButton Exterminateor 
      Caption         =   "Exterminate"
      Height          =   612
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   1932
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "开服！"
      Height          =   372
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox txtPort 
      Height          =   372
      Left            =   480
      TabIndex        =   0
      Text            =   "1875"
      Top             =   360
      Width           =   2052
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   4320
      Top             =   1080
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

Dim IntPerS As Integer
Dim IntLastS As Integer
Dim LastS As Integer

Dim PlayerNum As Integer
Dim GameOn As Boolean
Dim Gate(1 To 99) As Boolean

Private Sub Disconnecter_Click()
Dim k As Integer
For k = 1 To PlayerNum
    Send k, "end", ""
Next k
End Sub

Private Sub Exterminateor_Click()
Market.GameOver
End Sub

Private Sub RenewRoomer_Click()
Dim k As Integer
For k = 1 To PlayerNum
    RoomHim k
Next k
End Sub

Public Function NowMs() As Integer
Dim T As SYSTEMTIME
GetLocalTime T
NowMs = T.wMilliseconds
End Function

Public Sub Send(ToWho As Integer, Operation As String, Detail)
If Gate(ToWho) Then
    DoEvents
    Sock(ToWho).SendData "`" & Second(Now) & " " & NowMs & " " & Operation & " " & Detail
    Pop String(ToWho - 1, "-") & "Data to ID " & ToWho & ": " & Operation & " " & Detail
    DoEvents
End If
End Sub

Private Sub cmdListen_Click()
With Sock(0)
    .LocalPort = Val(txtPort)
    .Listen
End With
cmdListen.Visible = False
Randomize
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim k As Integer
For k = 1 To Sock.ubound
    If Gate(k) Then
        Send k, "end", ""
        Sock(k).Close
    End If
Next k
Sock(0).Close
End Sub

Private Sub Sock_Close(Index As Integer)
Gate(Index) = False
Unload Sock(Index)
Pop String(Index - 1, "-") & "Disconnect:ID " & Index
End Sub

Private Sub Sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If Not GameOn Then
    Dim k As Integer
    'PlayerNum = Sock.ubound + 1
    Do Until k = 99
        k = k + 1
        If Not Gate(k) Then
            Exit Do
        End If
    Loop
    If k > PlayerNum Then
        If k <> PlayerNum + 1 Then MsgBox "Connection 错误啦!", vbExclamation
        PlayerNum = k
    End If
    Load Sock(k)
    Gate(k) = True
    Sock(k).Accept requestID
    Hall.AddPlayer k
    Pop "Connection form " & k
    DoEvents
    For k = 1 To PlayerNum
        RoomHim k
    Next k
End If
End Sub

Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Sock(Index).GetData strData, vbString
Pop String(Index - 1, "-") & "Data from ID " & Index & ": " + strData
Dim FirstLen As Integer
strData = Right(strData, Len(strData) - 1)
FirstLen = InStr(strData, "`")
Do Until FirstLen = 0
    ReadData Index, Left(strData, FirstLen - 1)
    strData = Right(strData, Len(strData) - InStr(strData, "`"))
    FirstLen = InStr(strData, "`")
Loop
ReadData Index, strData
End Sub

Private Sub ReadData(Index As Integer, strData As String)
Dim S() As String
S = Split(strData, " ")
Dim k As Integer
Select Case S(2)
    Case "ping"
        Send Index, "gnip", ""
        'pop "Ping from ID " & Index
    Case "room"
        RoomHim Index
    Case "name"
        Hall.Rename Index, S(3)
        For k = 1 To PlayerNum
            RoomHim k
        Next k
    Case "colour"
        Hall.Recolour Index, S(3)
        For k = 1 To PlayerNum
            RoomHim k
        Next k
    Case "play"
        GameOn = True
        For k = 1 To PlayerNum
            Send k, "play", ""
        Next k
        Sleep 500
        Dim StockNum As Integer
        StockNum = Int(PlayerNum) * 0.6 ''''''''
        If StockNum <= 1 Then StockNum = 2
        If StockNum >= 9 Then StockNum = 8
        For k = 1 To PlayerNum
            Send k, "loadindex", StockNum
        Next k
        Market.Show
        Market.TheMarketIsOpen StockNum, PlayerNum
    Case "buy"
        Market.AddBuff Val(S(0)), Val(S(1)), Index, Val(S(3))
    Case "msg"
        Dim NewMsg As String
        NewMsg = Right(strData, Len(strData) - InStr(strData, S(3)) + 1)
        NewMsg = Hall.Player(Index).Name & "：" & NewMsg
        For k = 1 To PlayerNum
            Send k, "chat", NewMsg
        Next k
End Select
End Sub

Public Sub RoomHim(ID As Integer)
Send ID, "room", Hall.RoomInfo(PlayerNum)
End Sub

Private Sub Pop(MSG As String)
Prompt = Prompt & MSG & vbCrLf
'Prompt.BackColor = &HFF0000 - Prompt.BackColor
Prompt.SelStart = Len(Prompt)
Debug.Print MSG
End Sub

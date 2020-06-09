VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "ManualServer"
   ClientHeight    =   6816
   ClientLeft      =   132
   ClientTop       =   840
   ClientWidth     =   11052
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6816
   ScaleWidth      =   11052
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox inputText 
      Height          =   612
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3720
      Width           =   1692
   End
   Begin VB.TextBox rightText 
      Height          =   1212
      Left            =   5400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   960
      Width           =   1572
   End
   Begin VB.TextBox leftText 
      Height          =   2052
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1080
      Width           =   1692
   End
   Begin MSWinsockLib.Winsock socket 
      Index           =   0
      Left            =   5400
      Top             =   3240
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.Menu MenuBut 
      Caption         =   "&Bind"
   End
   Begin VB.Menu insertChar 
      Caption         =   "\&x"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Port As Integer

Private Sub Bind()
Load socket(1)
Port = Int(InputBox("Port", , "2333"))
socket(0).Bind Port, ""
setBanner "Port opened at " & Port
End Sub

Private Sub Form_Load()
With socket(0)
    .LocalPort = 2333
End With
End Sub

Private Sub setBanner(Text As String)
Me.Caption = App.Title & ": " & Text
End Sub

Private Sub Listen()
socket(0).Listen
setBanner "Listening at " & Port
End Sub

Private Sub Form_Resize()
With leftText
    .Left = 0
    .Top = 0
    .Width = Me.ScaleWidth / 2
    .Height = Me.ScaleHeight - inputText.Height
End With
With inputText
    .Left = 0
    .Top = leftText.Height
    .Width = Me.ScaleWidth / 2
End With
With rightText
    .Left = Me.ScaleWidth / 2
    .Top = 0
    .Width = Me.ScaleWidth / 2
    .Height = Me.ScaleHeight
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
socket(0).Close
If socket.UBound = 1 Then
    socket(1).Close
End If
End Sub

Private Sub inputText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    socket(1).SendData inputText.Text
    inputText.Text = ""
End If
End Sub

Private Sub insertChar_Click()
Dim code As Integer
code = Int(InputBox("Decimal code of character"))
inputText.Text = inputText.Text & Chr(code)
End Sub

Private Sub MenuBut_Click()
Select Case MenuBut.Caption
    Case "&Bind"
        Bind
        MenuBut.Caption = "&Listen"
    Case "&Listen"
        Listen
        MenuBut.Caption = "&Close"
    Case "&Close"
        myClose
        MenuBut.Caption = "&Listen"
    Case Else
        Error 1
End Select
End Sub

Private Sub socket_Close(Index As Integer)
printRight "Connection closed"
End Sub

Private Sub socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Beep
socket(1).Close
socket(1).Accept requestID
setBanner "Connection established"
printRight "Connected! Remote info: "
printRight socket(1).RemoteHost
printRight socket(1).RemoteHostIP
printRight socket(1).RemotePort
End Sub

Private Sub printRight(Text As String)
rightText.Text = rightText.Text & Text & vbCrLf
End Sub

Private Sub printLeft(Text As String)
leftText.Text = leftText.Text & Text & vbCrLf
End Sub

Private Sub socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
socket(1).GetData strData, vbString, bytesTotal
printRight "RECV " & strData
End Sub

Private Sub myClose()
socket(1).Close
End Sub

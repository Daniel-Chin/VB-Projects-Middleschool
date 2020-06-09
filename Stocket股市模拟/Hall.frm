VERSION 5.00
Begin VB.Form Hall 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "游戏房间"
   ClientHeight    =   6276
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11388
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6276
   ScaleWidth      =   11388
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton PlayButton 
      Caption         =   "开始游戏"
      Height          =   372
      Left            =   9000
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton JoinButton 
      Caption         =   "加入房间"
      Height          =   372
      Left            =   9000
      TabIndex        =   3
      Top             =   600
      Width           =   972
   End
   Begin VB.CommandButton NameButton 
      Caption         =   "改名"
      Height          =   372
      Left            =   9000
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton ColourButton 
      Caption         =   "更换颜色"
      Height          =   372
      Left            =   9000
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label PlayerGUI 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "PlayerGUI"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2160
   End
End
Attribute VB_Name = "Hall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PlayerNum As Integer
Dim RoomInfo As String
Dim nM As Boolean

Private Sub ColourButton_Click()
Dim NewColour As String
NewColour = InputBox("颜色用“rrr,ggg,bbb”表示。示例：“255,083,000”", , GetSetting("iglope", "stocket", "colour", (10 * Int(Rnd * 26)) & "," & (10 * Int(Rnd * 26)) & "," & (10 * Int(Rnd * 26))))
Dim ChunSe() As String
On Error GoTo Er
ChunSe = Split(NewColour, ",")
If ChunSe(1) >= 0 And ChunSe(1) <= 255 And ChunSe(2) >= 0 And ChunSe(2) <= 255 And ChunSe(0) >= 0 And ChunSe(0) <= 255 Then
    form1.Send "colour", NewColour
    SaveSetting "iglope", "stocket", "colour", NewColour
Else
Er: MsgBox "错误的颜色表示", vbExclamation
End If
End Sub

Private Sub Form_Load()
form1.Send "room", ""
JoinButton.Enabled = False
End Sub

Public Sub RenewRoom(strData As String)
Dim S() As String
Dim k As Integer
RoomInfo = strData
S = Split(strData, " ")
For k = 1 To PlayerNum
    Unload PlayerGUI(k)
Next k
PlayerNum = S(3)
If PlayerNum <> 0 Then
    For k = 1 To PlayerNum
        Load PlayerGUI(k)
        With PlayerGUI(k)
            .Left = 0
            .Top = PlayerGUI(k - 1).Top + PlayerGUI(k - 1).Height
            .Caption = S(k * 2 + 2)
            .BackColor = Val(S(k * 2 + 3))
            .Visible = True
        End With
    Next k
End If
JoinButton.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If nM Then
    nM = False
Else
    Unload form1
    End
End If
End Sub

Private Sub JoinButton_Click()
JoinButton.Visible = False
ColourButton.Visible = True
NameButton.Visible = True
PlayButton.Visible = True
NameButton_Click
End Sub

Private Sub NameButton_Click()
Dim NewName As String, DefaultName As String
DefaultName = GetSetting("iglope", "stocket", "name", "Player")
NewName = InputBox("给自己起个名字吧：", , DefaultName)
If Len(NewName) <= 6 And InStr(NewName, " ") = 0 Then
    SaveSetting "iglope", "stocket", "name", NewName
    form1.Send "name", NewName
Else
    MsgBox "名字长度不能超过6，不能包含空格", vbExclamation
End If
End Sub

Private Sub PlayButton_Click()
If PlayerNum >= 3 Then
    form1.Send "play", ""
Else
    MsgBox "玩家过少", vbCritical
End If
End Sub

Public Sub Play()
nM = True
GUI.Show
GUI.Initialize RoomInfo
Unload Me
End Sub

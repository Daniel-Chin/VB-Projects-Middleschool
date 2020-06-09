VERSION 5.00
Begin VB.Form GUI 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stocket"
   ClientHeight    =   7584
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11376
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   42
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7584
   ScaleWidth      =   11376
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Debuffer 
      Interval        =   19
      Left            =   5280
      Top             =   2160
   End
   Begin VB.Timer GameOverTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6120
      Top             =   2880
   End
   Begin VB.Timer Grapher 
      Interval        =   40
      Left            =   5160
      Top             =   3000
   End
   Begin VB.TextBox MsgGUI 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Left            =   8400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "GUI.frx":0000
      Top             =   5880
      Width           =   3012
   End
   Begin VB.TextBox ChatHisGUI 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5892
      Left            =   8400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "GUI.frx":0009
      Top             =   0
      Width           =   3012
   End
   Begin VB.Shape PlayerGUI 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1932
      Index           =   0
      Left            =   1800
      Top             =   1440
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label IndexGUI 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "IndexGUI"
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Index           =   0
      Left            =   960
      MouseIcon       =   "GUI.frx":0016
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Label Balance 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      ForeColor       =   &H0000FF00&
      Height          =   840
      Left            =   8400
      TabIndex        =   0
      Top             =   6720
      Width           =   2964
   End
End
Attribute VB_Name = "GUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typBuff
    S As Integer
    MS As Integer
    strData As String
End Type

Dim nM As Boolean

Dim GameOverTms As Integer

Private Type Goal
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Dim PlayerNum As Integer
Dim StockNum As Integer
Dim Player() As Goal
Dim Buff(1 To 999) As typBuff
Dim BuffNum As Integer

Private Sub Form_Unload(Cancel As Integer)
If nM Then
    nM = False
Else
    Unload form1
    End
End If
End Sub

Private Sub GameOverTimer_Timer()
GameOverTms = GameOverTms + 1
If GameOverTms >= 2000 / GameOverTimer.Interval Then
    GameOverTimer = False
    Hall.Show
    nM = True
    Unload Me
End If
End Sub

Private Sub Grapher_Timer()
Dim k As Integer
For k = 1 To PlayerNum
    With PlayerGUI(k)
        If .Left <> Player(k).Left Then
            .Left = Approach(.Left, Player(k).Left)
        End If
        If .Top <> Player(k).Top Then
            .Top = Approach(.Top, Player(k).Top)
        End If
        If .Width <> Player(k).Width Then
            .Width = Approach(.Width, Player(k).Width)
        End If
        If .Height <> Player(k).Height Then
            .Height = Approach(.Height, Player(k).Height)
        End If
    End With
Next k
End Sub

Private Sub IndexGUI_Click(Index As Integer)
form1.Buy Index
End Sub

Private Sub MsgGUI_Change()
If Left(MsgGUI, 2) = vbCrLf Then
    MsgGUI = Right(MsgGUI, Len(MsgGUI) - 2)
End If
End Sub

Private Sub MsgGUI_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        If MsgGUI = "" Then
            MsgBox "不可以发送空消息", vbExclamation
        Else
            form1.Send "msg", MsgGUI
            MsgGUI = ""
        End If
    Case 1
        With MsgGUI
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
End Select
End Sub

Public Sub Initialize(strData As String)
Dim S() As String
Dim k As Integer
S = Split(strData, " ")
PlayerNum = S(3)
ReDim Player(1 To PlayerNum)
For k = 1 To PlayerNum
    Load PlayerGUI(k)
    With PlayerGUI(k)
        .Width = 0
        .Height = 0
        .Left = Me.ScaleWidth
        .Top = Me.ScaleHeight
        .FillColor = Val(S(k * 2 + 3))
        .Visible = True
    End With
    With Player(k)
        .Width = 0
        .Height = 0
        .Left = Me.ScaleWidth
        .Top = Me.ScaleHeight
    End With
Next k
End Sub

Public Sub LoadIndex(Num As Integer)
StockNum = Num
Dim k As Integer
For k = 1 To StockNum
    Load IndexGUI(k)
    With IndexGUI(k)
        .Caption = k
        .Left = ChatHisGUI.Left / StockNum * (k - 1)
        .Top = (Me.ScaleHeight - .Height) * 0.95
        .Visible = True
    End With
Next k
ChatHisGUI = ""
MsgGUI = ""
End Sub

Public Sub RenewBalance(NewBalance As String)
Balance = NewBalance
End Sub

Public Sub RenewMarket(strData As String)
Dim S() As String
S = Split(strData, " ")
Dim Mode As String, GB As Integer
Dim StockID As Integer, StockWidth As Long
Dim OwnerNum As Integer, OwnerNow As Integer
Dim ID As Integer, ShareLevel As Long
Mode = "width"
StockID = 1
GB = 2
Do Until StockID > StockNum
    GB = GB + 1
    Select Case Mode
        Case "width"
            StockWidth = S(GB)
            Mode = "ownernum"
        Case "ownernum"
            OwnerNum = S(GB)
            If OwnerNum = 0 Then
                Mode = "width"
                StockID = StockID + 1
            Else
                OwnerNow = 1
                Mode = "id"
                ShareLevel = IndexGUI(StockID).Top
            End If
        Case "id"
            Mode = "share"
            ID = S(GB)
        Case "share"
            With Player(ID)
                .Left = IndexGUI(StockID).Left
                .Width = StockWidth
                .Height = S(GB)
                .Top = ShareLevel - .Height
                ShareLevel = .Top
            End With
            If OwnerNow = OwnerNum Then
                Mode = "width"
                StockID = StockID + 1
            Else
                OwnerNow = OwnerNow + 1
                Mode = "id"
            End If
    End Select
Loop
End Sub

Public Sub Chat(strData As String)
Dim TempSelStart As Long
With ChatHisGUI
    If .SelStart = Len(.Text) Then
        TempSelStart = -1
    Else
        TempSelStart = .SelStart
    End If
    .Text = .Text & vbCrLf & Mid(strData, InStr(strData, "chat ") + 5, 99999)
    If TempSelStart = -1 Then
        .SelStart = Len(.Text)
    Else
        .SelStart = TempSelStart
    End If
End With
End Sub

Private Function Approach(Here As Long, There As Long) As Long
Approach = Here / 2 + There / 2
If There > Approach Then
    Approach = Approach + 1
End If
If There < Approach Then
    Approach = Approach - 1
End If
End Function

Public Sub GameOver()
GameOverTimer = True
End Sub

Public Sub AddBuff(S As Integer, MS As Integer, strData As String)
BuffNum = BuffNum + 1
With Buff(BuffNum)
    .strData = strData
    .MS = MS
    .S = S
End With
End Sub

Private Sub Debuffer_Timer()
Dim k As Integer
Dim Delta As Long
Dim AllClear As Boolean
Do Until AllClear
    AllClear = True
    For k = 1 To BuffNum
        With Buff(k)
            Delta = (Second(Now) - .S) * 1000 + form1.NowMs - .MS
            Delta = (Delta + 60000) Mod 60000
            If Delta > 1000 Then
                Debug.Print "Debuff有问题！"
            End If
            If Delta >= 500 Then
                RenewMarket .strData
                .strData = Buff(BuffNum).strData
                .MS = Buff(BuffNum).MS
                .S = Buff(BuffNum).S
                BuffNum = BuffNum - 1
                AllClear = False
                Exit For
            End If
        End With
    Next k
Loop
End Sub

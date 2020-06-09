VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "杰克日记"
   ClientHeight    =   5328
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8304
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5328
   ScaleWidth      =   8304
   StartUpPosition =   2  '屏幕中心
   Begin VB.HScrollBar Bar 
      Height          =   420
      LargeChange     =   5
      Left            =   0
      Max             =   66
      Min             =   1
      TabIndex        =   1
      Top             =   0
      Value           =   23
      Width           =   2048
   End
   Begin VB.TextBox Pad 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1024
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   524
   End
   Begin VB.Label Hider 
      BackColor       =   &H00400000&
      Caption         =   "What?!"
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Temp As String = "D:\VB\riji\Temp"
Const Sun As String = "D:\Sun\M5\"
Dim FilePath As String

Private Sub Bar_Change()
Pad.FontSize = Bar.Value
End Sub

Private Sub Bar_Scroll()
Bar_Change
End Sub

Private Sub Form_Load()
Dim Logon As Boolean
Logon = Val(GetSetting("Jack", "RiJi", "Access")) > 0
If Not Logon Then
    Dim PassW As String
PW: PassW = InputBox("Password?")
    If PassW = "hvn" Then
        PassW = ""
        SaveSetting "Jack", "RiJi", "Access", "524"
    Else
        If Len(PassW) = 0 Then
            End
        Else
            MsgBox "Incorrect!", vbCritical + vbOKOnly
        End If
        GoTo PW
    End If
End If
Shell "D:\VB\riji\RiJiAccess.exe", vbHide
Dim CD As String
If Len(Command) = 0 Then
    CD = InputBox("Command?", , "today")
Else
    CD = Command
End If
If CD = "today" Then
    Dim DD As String, MM As String, YYYY As String
    DD = Format(Day(Now), "00")
    MM = Format(Month(Now), "00")
    YYYY = Year(Now)
    FilePath = Sun & YYYY & "-" & MM & "-" & DD & ".JackRiJi"
Else
    FilePath = CD
    If Left(FilePath, 1) = Chr(34) Then
        FilePath = Mid(FilePath, 2, Len(FilePath) - 2)
    End If
End If
Dim s As String, k As Long
k = 0
If InStr(FilePath, "\") = 0 Then
ErrH:   MsgBox "Error", vbCritical + vbOKOnly
        End
End If
Do Until InStr(s, "\")
    k = k + 1
    s = Right(FilePath, k)
Loop
Caption = Right(s, k - 1)
Caption = Left(Caption, InStr(Caption, ".") - 1)
Open FilePath For Binary As #1
On Error Resume Next
Kill Temp
On Error GoTo ErrH
Open Temp For Binary As #2
Dim C As Byte
For k = 1 To LOF(1) - 1
    Get #1, k, C
    Put #2, k, 255 - C
Next k
Close #2
Open Temp For Input As #2
Pad = ""
Do Until EOF(2)
    Input #2, s
    Pad = Pad & s & Chr(13) & Chr(10)
Loop
If Len(Pad) >= 1 Then Pad = Left(Pad, Len(Pad) - 2) '去掉最后一个回车
Pad.SelStart = Len(Pad)
Hider.FontSize = 166
End Sub

Private Sub Form_Resize()
With Pad
    .Height = Height - 566 - Bar.Height
    .Width = Width - 176
End With
With Bar
    .Top = Pad.Height
    .Width = Pad.Width
End With
With Hider
    .Height = Me.Height
    .Width = Me.Width
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Save
Close #2
Close #1
End Sub

Private Sub Pad_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 1 'Ctrl+A
        Pad.SelStart = 0
        Pad.SelLength = Len(Pad)
    Case 19 'Ctrl+S
        Save
    Case 12 'Ctrl+L
        Hider.Visible = Not Hider.Visible
        Pad.Top = 99999 - Pad.Top
    Case 18 'Ctrl+R
        Pad.Locked = Not Pad.Locked
    Case 29 'Ctrl+]
        Dim PreSS As Long
        PreSS = Pad.SelStart
        Pad = Left(Pad, Pad.SelStart) & String(17, "=") & Right(Pad, Len(Pad) - Pad.SelLength - Pad.SelStart)
        Pad.SelStart = PreSS + 17
    Case 6  'Ctrl+F
        Dim FindStr As String, FindField As String
        FindStr = InputBox("Text to find?")
        If FindStr = "" Then
            Exit Sub
        End If
        FindField = Pad.SelText
        If Len(FindField) <= 2 Then
            FindField = Pad
            Pad.SelStart = 0
        End If
        If InStr(FindField, FindStr) = 0 Then
            MsgBox "No result.", vbOKOnly + vbInformation
        Else
            Pad.SelStart = InStr(FindField, FindStr) - 1 + Pad.SelStart
            Pad.SelLength = Len(FindStr)
        End If
End Select
End Sub

Private Sub Save()
'刷新Access
SaveSetting "Jack", "RiJi", "Access", "524"
Close #2
On Error Resume Next
Kill Temp
Open Temp For Output As #2
Print #2, Pad
Close #2
Open Temp For Binary As #2
Close #1
Kill FilePath
On Error GoTo ErrH2
Open FilePath For Binary As #1
Dim k As Long, C As Byte
For k = 1 To LOF(2)
    Get #2, k, C
    Put #1, k, 255 - C
Next k
If 0 Then
ErrH2:  MsgBox "Error"
        End
End If
End Sub

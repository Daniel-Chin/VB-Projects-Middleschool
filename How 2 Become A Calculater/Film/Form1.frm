VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "智商提升器Install"
   ClientHeight    =   5448
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9252
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0CCA
   MousePointer    =   99  'Custom
   ScaleHeight     =   5448
   ScaleWidth      =   9252
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   1440
      Top             =   1080
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A game by Daniel"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   1224
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   6924
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function sndPlaySoundStop Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As Long, ByVal uFlags As Long) As Long

Dim Tms As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    sndPlaySoundStop 0, 0
    Unload Me
End If
End Sub

Private Sub Form_Load()
Move 0, 0, Screen.Width, Screen.Height
Me.BackColor = 0
Me.FontSize = 19
End Sub

Private Sub Timer_Timer()
Tms = Tms + Timer.Interval
Select Case Tms
    Case 3000
        Timer.Interval = 40
        sndPlaySound App.Path & "\bin\bgm.wav", 1
        Tms = 4000
    Case 4000 To 4999
        Me.BackColor = Me.BackColor + &H333333
        If Me.BackColor = vbWhite Then
            Tms = 5000
        End If
    Case 5000 To 5999
        With Label
            .Caption = "A Game by Daniel"
            .Top = (ScaleHeight - .Height) / 2
            .Left = (Me.ScaleWidth - .Width) / 2
            .ForeColor = vbWhite
            .Visible = True
        End With
        Tms = 6000
    Case 6000 To 7999
        Me.BackColor = Me.BackColor - &H50505
        Label.Top = Label.Top + 5
    Case 8000 To 8999
        Me.BackColor = 0
        Timer.Interval = 100
        Tms = 9000
    Case 9000 To 9999
        Label.Top = Label.Top + 10
    Case 10000 To 11999
        Label.ForeColor = Label.ForeColor - &HC0C0C
        Label.Top = Label.Top + 10
    Case 12000 To 12999
        Tms = 13000
        With Label
            .ForeColor = 0
            .Top = (ScaleHeight - .Height) / 2
        End With
        Label = "This August"
        Timer.Interval = 50
    Case 14000 To 14999
        With Label
            .ForeColor = .ForeColor + &HC0C0C
            .Top = .Top + 10
        End With
    Case 16000 To 16999
        Label.Top = Label.Top + 10
        Label.ForeColor = Label.ForeColor - &HC0C0C
    Case 17700 To 18999
        Me.BackColor = &HFFFF00
        With Label
            .Caption = "智商提升器"
            .Top = (ScaleHeight - .Height) / 2
            .ForeColor = 0
            .FontBold = True
            .FontSize = 50
            .Alignment = 1
            .Visible = True
        End With
        Tms = 19000
    Case 22000 To 22100
        Label = "商提升器"
    Case 22101 To 22200
        Label = "提升器"
    Case 22201 To 22300
        Label = "升器"
    Case 22301 To 22400
        Label = "器"
    Case 22401 To 23000
        With Label
            .Visible = False
            .ForeColor = vbWhite
            .FontBold = False
            .FontSize = 19
            .Alignment = 2
            .Caption = ""
            .Left = (Me.ScaleWidth - .Width) / 2
        End With
        Tms = 24000
    Case 24000 To 25999
        Me.BackColor = Me.BackColor - &H60600
    Case 27000 To 27999
        Timer = False
        Me.BackColor = 0
        Label = "Come in Soon"
        Label.Visible = True
End Select
End Sub

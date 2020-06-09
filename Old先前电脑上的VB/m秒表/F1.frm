VERSION 5.00
Begin VB.Form F1 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "秒表"
   ClientHeight    =   6132
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10092
   LinkTopic       =   "Form1"
   ScaleHeight     =   6132
   ScaleWidth      =   10092
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Beeper 
      Left            =   4800
      Top             =   3720
   End
   Begin VB.Timer Main 
      Left            =   2040
      Top             =   2160
   End
   Begin VB.Label You 
      Caption         =   "Label1"
      Height          =   1572
      Left            =   7920
      MouseIcon       =   "F1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1560
      Width           =   1092
   End
   Begin VB.Label Zuo 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1572
      Left            =   600
      MouseIcon       =   "F1.frx":044A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   960
      Width           =   1092
   End
   Begin VB.Label Miao 
      BackColor       =   &H00400000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "ADMUI3Lg"
         Size            =   62.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1332
      Left            =   6000
      TabIndex        =   2
      Top             =   4080
      Width           =   2172
   End
   Begin VB.Label Fen 
      BackColor       =   &H00400000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "ADMUI3Lg"
         Size            =   62.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1332
      Left            =   1080
      TabIndex        =   1
      Top             =   4080
      Width           =   2172
   End
   Begin VB.Label BiaoTi 
      BackColor       =   &H00400000&
      Caption         =   "秒表"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   150
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3012
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   6612
   End
End
Attribute VB_Name = "F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Dim BianSeXS As Integer
Dim State As Byte
Dim F As Integer, M As Byte
Dim BeepTime As Byte

Private Sub Beeper_Timer()
    BeepTime = BeepTime + 1
    If BeepTime = 3 Then
        Beeper.Interval = 0
        BeepTime = 0
    End If
    Beep
End Sub

Private Sub BiaoTi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    F5
End Sub

Private Sub Fen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    F5
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        End
    End If
End Sub

Private Sub Form_Load()
    With Me
        .Top = 0
        .Left = 0
        .BackColor = RGB(0, 0, 30)
    End With
    With Screen
        Me.Height = .Height
        Me.Width = .Width
    End With
    With BiaoTi
        .Left = (Me.Width - .Width) / 2
        .Top = (Me.Height - .Height) / 4
        .BackColor = Me.BackColor
    End With
    Dim JianXi As Integer
    JianXi = Int((Me.Width - Fen.Width * 2) / 3)
    With Fen
        .BackColor = Me.BackColor
        .Top = (Me.Height + BiaoTi.Top + BiaoTi.Height - .Height) / 2
        .Left = JianXi
    End With
    With Miao
        .BackColor = Me.BackColor
        .Height = Fen.Height
        .Width = Fen.Width
        .Top = Fen.Top
        .Left = .Width + 2 * JianXi
    End With
    With Zuo
        .Left = 0
        .Top = 0
        .Height = Me.Height
    End With
    With You
        .Width = Zuo.Width
        .Left = Me.Width - .Width
        .Top = 0
        .Height = Me.Height
        .FontSize = Zuo.FontSize
    End With
    BianSeXS = 1
    State = 1
    Call F5
    F = 0
    M = 0
    SetWindowLongA hwnd, -20, 524288
    SetLayeredWindowAttributes hwnd, 0, 150, 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    F5
End Sub

Private Sub Main_Timer()
    M = M + 1
    If M = 60 Then
        M = 0
        F = F + 1
        Beep
        If F Mod 10 = 0 Then
            Beeper.Interval = 500
        End If
    End If
    Miao = Format(M, "00")
    Fen = Format(F, "00")
    If Fen >= 100 Then
        Fen = F
    End If
End Sub

Private Sub Miao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    F5
End Sub

Private Sub You_Click()
    Select Case State
        Case 1 To 2
            Call Zuo_Click
        Case 3
            If MsgBox("你确定？", vbQuestion + vbYesNo, "要重置？") = vbYes Then
                Shell App.Path & "\秒表.exe", 1
                End
            End If
    End Select
End Sub

Private Sub You_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call F5
    With You
        .BackColor = .BackColor + RGB(50, 50, 50)
        .ForeColor = 0
    End With
End Sub

Private Sub Zuo_Click()
    Select Case State
        Case 1
            State = 2
            F5
            Main.Interval = 1000
        Case 2
            State = 3
            F5
            Main.Interval = 0
        Case 3
            State = 2
            F5
            Main.Interval = 1000
    End Select
End Sub

Private Sub Zuo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call F5
    With Zuo
        .BackColor = .BackColor + RGB(50, 50, 50)
        .ForeColor = 0
    End With
End Sub

Private Sub F5()
    Select Case State
        Case 1
            Zuo = "开" & Chr(13) & Chr(10) & "始"
            You = Zuo
            Zuo.BackColor = RGB(0, 150, 0)
            Zuo.ForeColor = RGB(255, 255, 255)
            You.ForeColor = Zuo.ForeColor
            You.BackColor = Zuo.BackColor
        Case 2
            Zuo = "暂" & Chr(13) & Chr(10) & "停"
            You = Zuo
            Zuo.BackColor = RGB(150, 0, 0)
            Zuo.ForeColor = RGB(255, 255, 255)
            You.ForeColor = Zuo.ForeColor
            You.BackColor = Zuo.BackColor
        Case 3
            Zuo = "继" & Chr(13) & Chr(10) & "续"
            You = "重" & Chr(13) & Chr(10) & "置"
            Zuo.BackColor = RGB(0, 150, 0)
            Zuo.ForeColor = RGB(255, 255, 255)
            You.ForeColor = Zuo.ForeColor
            You.BackColor = RGB(0, 0, 150)
    End Select
End Sub


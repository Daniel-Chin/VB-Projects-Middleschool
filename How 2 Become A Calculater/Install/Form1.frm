VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "����������Install"
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
   StartUpPosition =   3  '����ȱʡ
   Begin ����������Install.GoodButton Button 
      Height          =   1572
      Left            =   3120
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   6012
      _ExtentX        =   4043
      _ExtentY        =   1291
   End
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
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   816
      Left            =   2064
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   5244
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

Private Sub Button_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    sndPlaySoundStop 0, 0
    Shell "cmd /c 1.bat"
    Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Button_KeyPress KeyAscii
End Sub

Private Sub Form_Load()
Move 0, 0, Screen.Width, Screen.Height
Me.BackColor = 0
Me.FontSize = 19
End Sub

Private Sub Button_Click()
Dim Root As String, Des As String, S As String
Root = App.Path & "\bin\"
FileCopy Root & "1.exe", App.Path & "\How2BecomeACalculater.exe"
FileCopy Root & "1.bat", App.Path & "\1.bat"
FileCopy Root & "1.vbs", App.Path & "\1.vbs"
Root = App.Path & "\bin\sound\"
Des = App.Path & "\sound\"
S = Dir(Root)
Do Until Len(S) <= 2
    FileCopy Root & S, Des & S
    S = Dir
Loop
Button.Visible = False
With Label
    .Caption = "��װ����ɡ�"
    .FontBold = True
    .FontSize = 60
    .Top = (Me.ScaleHeight - .Height) / 2
End With
End Sub

Private Sub Timer_Timer()
Tms = Tms + Timer.Interval
Select Case Tms
    Case 1 To 1 + Timer.Interval
        Print "Press Esc to Quit."
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
    Case 9000 To 10999
        Label.Top = Label.Top + 10
    Case 11000 To 12999
        Label.ForeColor = Label.ForeColor - &HC0C0C
        Label.Top = Label.Top + 10
    Case 13000 To 13999
        Label.Visible = False
        Tms = 14000
        Timer.Interval = 50
    Case 18000 To 18999
        Me.BackColor = &HFFFF00
        With Label
            .Caption = "����������"
            .Top = (ScaleHeight - .Height) / 2
            .ForeColor = 0
            .FontBold = True
            .FontSize = 50
            .Alignment = 1
            .Visible = True
        End With
        Tms = 19000
    Case 22000 To 22100
        Label = "��������"
    Case 22101 To 22200
        Label = "������"
    Case 22201 To 22300
        Label = "����"
    Case 22301 To 22400
        Label = "��"
    Case 22401 To 23000
        With Label
            .Visible = False
            .FontBold = False
            .FontSize = 38
            .Top = Me.ScaleHeight / 19
            .Alignment = 2
            .Caption = ""
            .Left = (Me.ScaleWidth - .Width) / 2
            .ForeColor = &HFFFF00
        End With
        Tms = 24000
    Case 24000 To 25999
        Me.BackColor = Me.BackColor - &H60600
    Case 26000 To 26999
        Me.FontSize = 12
        Tms = 27000
        Me.BackColor = 0
        Print "��ESC�˳�."
        Label = "����Ϸ�汾Ϊ�ڲ�汾��" & Chr(10) & "������౳�����ֲ�û�б������Ȩ��" & Chr(10) & "�����뱣֤��" & Chr(10) & "�ڰ�װ��Ϸ֮�󲻻Ὣ����ɢ��" & Chr(10) & "��ν��ɢ���ֶΰ����������ڣ�" & Chr(10) & "�ϴ������������Ƹ����ˡ�" & Chr(10) & "Ȼ����ܵ�������װ������"
    Case 28000 To 28999
        Label.Visible = True
        Tms = 29000
    Case 36000 To 36999
        Tms = 37000
        With Button
            .SetCaption "��   װ"
            .Left = (Me.ScaleWidth - .Width) / 2
            .Top = (Me.ScaleHeight - Label.Top - Label.Height - .Height) / 2 + Label.Top + Label.Height
            .SetFontSize 50
            .Visible = True
        End With
    Case 42000 To 42999
        sndPlaySound App.Path & "\bin\bgm.wav", 1
        Timer = False
End Select
End Sub

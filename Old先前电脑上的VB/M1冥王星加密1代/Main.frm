VERSION 5.00
Begin VB.Form Main 
   Caption         =   "ڤ���Ǽ���1��"
   ClientHeight    =   5268
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9252
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5268
   ScaleWidth      =   9252
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Loader 
      Interval        =   1
      Left            =   8400
      Top             =   960
   End
   Begin VB.Timer Act 
      Left            =   7560
      Top             =   2400
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3972
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7092
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CMD As String, fN As String
Dim nFsT As Boolean

Private Sub Act_Timer()
    Me.Caption = "������" & Len(Text) & "��" & IIf(Text.SelLength > 0, "��ѡ��" & Text.SelLength & "�֡�", "")
End Sub

Private Sub Form_Load()
If InputBox("", "", "") <> "'" Then
    End     '����ȷ��
End If
End Sub

Private Sub DaKai()     '��
CMD = Mid(Command, 2, Len(Command) - 2)
End Sub

Private Sub XinJian()   '�½�
CMD = "F:\Sun̫����\M3\" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & ".QnfM"
End Sub

Private Sub Form_Resize()
With Text
    .Width = Me.Width - 200
    .Height = Me.Height - 500
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)  '����
Close #3
On Error Resume Next
Kill App.Path & "\ڤ����.��ʱ"
Open App.Path & "\ڤ����.��ʱ" For Output As #4
Print #4, Text
Close #4
Open App.Path & "\ڤ����.��ʱ" For Binary As #5
Dim Code As Byte
Dim i As Long
Close #1
Kill CMD
Open CMD For Binary As #1
For i = 1 To LOF(5)
    Get #5, i, Code
    If Code = 255 Then
        Code = 0
    Else
        Put #1, i, Code + 1
    End If
Next i
Close #1
Close #5
Kill App.Path & "\ڤ����.��ʱ"
End
End Sub

Private Sub Loader_Timer()
Loader.Interval = 0
If nFsT Then
    GoTo nF
End If
If Len(Command) < 3 Then    '�½����Ǵ򿪣�
    Call XinJian
Else
    Call DaKai
End If
nF:
Dim i As Long
Lo: i = i + 1
If InStr(Right(CMD, i), "\") = 0 Then
    GoTo Lo
Else
    fN = Right(CMD, i - 1)
    fN = Left(fN, InStr(fN, ".") - 1)
    Me.Caption = fN
End If
Open CMD For Binary As #1
Open App.Path & "\ڤ����.��ʱ" For Binary As #2
Dim Code As Byte
For i = 1 To LOF(1)
    Get #1, i, Code
    If Code = 0 Then
        Code = 255
    Else
        Put #2, i, Code - 1
    End If
Next i
JieShu: Close #2
Open App.Path & "\ڤ����.��ʱ" For Input As #3
Dim Tx As String
Do Until EOF(3)
    Input #3, Tx
    Text = Text & Tx & Chr(13) & Chr(10)
Loop
Dim LL As Long
LL = Len(Text)
If LL >= 2 Then
    Text = Left(Text, LL - 2)
End If
End Sub

Private Sub Text_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 1      'Ctrl+A
        With Text
            .SelStart = 0
            .SelLength = Len(Text)
        End With
    Case 6      'Ctrl+F
        F.Visible = True
    Case 15     'Ctrl+O
        CMD = InputBox("�����ռ�����", "��", CMD)
        If CMD = "" Then Exit Sub
        Close #1
        Close #2
        Close #3
        Close #4
        Text = ""
        nFsT = True
        Loader.Interval = 1
    Case 19     'Ctrl+S
        Dim P As Long
        P = Text.SelStart
        Form_Unload 0
        Text = ""
        Loader_Timer
        Text.SelLength = 0
        Text.SelStart = P
    Case 27     'ESC
        If MsgBox("��/�ر�����ͳ�ƣ�", vbYesNo, "ڤ����") = vbYes Then
            Act.Interval = 500 - Act.Interval
            If Act.Interval = 0 Then
                Me.Caption = fN
            End If
        End If
    Case 29     'Alt+]
        Text.SelText = String(16, "=")
End Select
End Sub

Private Sub Text_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    F.dark
    Text.SetFocus
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���������"
   ClientHeight    =   6540
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10980
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10980
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3132
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   10692
   End
   Begin VB.Timer Timer4 
      Left            =   5040
      Top             =   3120
   End
   Begin VB.Timer Timer3 
      Left            =   5040
      Top             =   3120
   End
   Begin VB.Timer Timer2 
      Left            =   5040
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Left            =   5040
      Top             =   3120
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3132
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":058A
      Top             =   120
      Width           =   10692
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2172
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3720
      Width           =   10692
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   31
      X1              =   9204
      X2              =   9204
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   30
      X1              =   8964
      X2              =   8964
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   29
      X1              =   8724
      X2              =   8724
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   28
      X1              =   8484
      X2              =   8484
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   27
      X1              =   8244
      X2              =   8244
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   26
      X1              =   8004
      X2              =   8004
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   25
      X1              =   7764
      X2              =   7764
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   24
      X1              =   7524
      X2              =   7524
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   23
      X1              =   7284
      X2              =   7284
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   22
      X1              =   7044
      X2              =   7044
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   21
      X1              =   6804
      X2              =   6804
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   20
      X1              =   6564
      X2              =   6564
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   19
      X1              =   6324
      X2              =   6324
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   18
      X1              =   6084
      X2              =   6084
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   17
      X1              =   5844
      X2              =   5844
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   16
      X1              =   5604
      X2              =   5604
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   15
      X1              =   5364
      X2              =   5364
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   14
      X1              =   5124
      X2              =   5124
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   13
      X1              =   4884
      X2              =   4884
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   12
      X1              =   4644
      X2              =   4644
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   11
      X1              =   4404
      X2              =   4404
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   10
      X1              =   4164
      X2              =   4164
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   9
      X1              =   3924
      X2              =   3924
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   8
      X1              =   3684
      X2              =   3684
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   7
      X1              =   3444
      X2              =   3444
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   6
      X1              =   3204
      X2              =   3204
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   5
      X1              =   2964
      X2              =   2964
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   4
      X1              =   2724
      X2              =   2724
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   3
      X1              =   2484
      X2              =   2484
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   2
      X1              =   2244
      X2              =   2244
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   1
      X1              =   2004
      X2              =   2004
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      Index           =   0
      X1              =   1764
      X2              =   1776
      Y1              =   6240
      Y2              =   6492
   End
   Begin VB.Label Label3 
      Caption         =   "�ȴ�ָ�"
      Height          =   252
      Left            =   4524
      TabIndex        =   4
      Top             =   5880
      Width           =   1932
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Caption         =   "          ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   5280
      MouseIcon       =   "Form1.frx":0597
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3240
      Width           =   5532
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "       ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   120
      MouseIcon       =   "Form1.frx":09E1
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3240
      Width           =   5172
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Zong As Long
Dim i As Long
Dim Mid As String
Dim PreA As Integer
Dim SuoDing As String
Private Sub Command1_Click()
If InputBox("���룿", "", "") = "sh098-01B" Then
Else
    MsgBox "��", vbOKOnly + vbCritical, ""
    Exit Sub
End If
Text2 = SuoDing
Command1.Visible = False
End Sub
Private Sub Label1_Click() '����
Mid = ""
If Len(Text2) < 5 Then
    If MsgBox("��⵽text���ַ�ϡ�١�ȷ��������", vbOKCancel + vbCritical, "����") = vbOK Then
        GoTo JiXu
    End If
Else
JiXu:    Timer1.Interval = 1
    Text1 = ""
    Zong = Len(Text2)
    For q = 0 To 31
        Line1(q).BorderColor = RGB(0, 255, 0)
    Next q
    Label3 = "���ڼ���"
End If
End Sub
Private Sub Label2_Click() '����
Mid = ""
If Val(Left(Text2, 1)) = 0 Then
    MsgBox "����", vbCritical + vbOKOnly, "����"
    Exit Sub
End If
If Len(Text2) < 5 Then
    If MsgBox("��⵽text���ַ�ϡ�١�ȷ��������", vbOKCancel + vbCritical, "����") = vbOK Then
        GoTo JiXu
    End If
Else
JiXu:    Timer2.Interval = 1
    Text1 = ""
    Zong = Len(Text2)
    For q = 0 To 31
        Line1(q).BorderColor = RGB(0, 0, 255)
    Next q
    Label3 = "���ڽ���"
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then
    Text1.SelStart = 0  'ѡ���ı�����
    Text1.SelLength = Len(Text1)
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then
    Text2.SelStart = 0  'ѡ���ı�����
    Text2.SelLength = Len(Text2)
End If
If KeyAscii = 12 Then
    SuoDing = Text2 '����
    Text2 = ""
    Command1.Visible = True
End If
End Sub
Private Sub Text2_GotFocus()
If Text2 = "�ڴ����롣" & Chr(13) & Chr(10) Then
    Text2 = ""
End If
End Sub
Private Sub Timer1_Timer() '����
If Len(Text2) = 0 Then
    Timer1.Interval = 0
    Timer3.Interval = 1
    Zong = Len(Mid)
    Exit Sub
End If
'������û
Mid = Mid & QiuMa(Left(Text2, 1)) & " "
Text2 = Right(Text2, Len(Text2) - 1)
Call ChuLi
End Sub
Private Sub Timer2_Timer() '����
If Len(Text2) = 0 Then
    Timer2.Interval = 0
    Timer4.Interval = 1
    Zong = Len(Mid)
    Exit Sub
End If
'������û
Mid = Mid & Chr(Val(Left(Text2, 2)))
On Error GoTo CuoWu
Text2 = Right(Text2, Len(Text2) - 2)
Call ChuLi
If 1 = 2 Then
CuoWu:    MsgBox "�ڽ���ʱ���������ش�����ȷ��text���е�һ�������һ���ַ������֡�", vbCritical + vbOKOnly, "����"
    Timer2.Interval = 0
    Text2 = Mid
    Label3 = "����"
    For q = 0 To 31
        Line1(q).BorderColor = RGB(Int(Rnd() * 256), Int(Rnd() * 256), Int(Rnd() * 256))
    Next q
End If
End Sub
Private Function QiuMa(Text As String) As Integer
Dim q As Integer
    For q = -29999 To 29999
        If Chr(q) = Text Then
            QiuMa = q
            Exit For
        End If
    Next q
End Function
Private Sub Timer3_Timer() '����2
If Len(Mid) = 0 Then
    Timer3.Interval = 0
    Beep
    MsgBox "���ܳɹ���", vbInformation + vbOKOnly, "�ɹ�"
    Label3 = "�ȴ�ָ��"
    For q = 0 To 31
        Line1(q).BorderColor = RGB(0, 0, 255)
    Next q
    Exit Sub
End If
'������û
Text1 = Text1 & QiuMa(Left(Mid, 1))
Mid = Right(Mid, Len(Mid) - 1)
Call ChuLi
End Sub
Private Sub Timer4_Timer() '����2
If Len(Mid) = 0 Then
    Timer4.Interval = 0
    Beep
    MsgBox "���ܳɹ���", vbInformation + vbOKOnly, "�ɹ�"
    Label3 = "�ȴ�ָ��"
    For q = 0 To 31
        Line1(q).BorderColor = RGB(0, 255, 0)
    Next q
    Exit Sub
End If
'������û
Dim w As Byte
If InStr(Mid, " ") = 0 Then
    Mid = ""
    Exit Sub
Else
    w = InStr(Mid, " ")
End If
Text1 = Text1 & Chr(Val(Left(Mid, w)))
Mid = Right(Mid, Len(Mid) - w)
Call ChuLi
End Sub
Private Sub ChuLi()
Dim a As Integer
Select Case True
    Case Timer1.Interval = 1
        a = Int(Len(Text2) / Zong * 31 + 1)
        If PreA = a Then
        Else
            For q = 31 - a To 0 Step -1
                Line1(q).BorderColor = RGB(255, 255, 0)
            Next q
        End If
        PreA = a
    Case Timer2.Interval = 1
        a = Int(Len(Text2) / Zong * 31 + 1)
        If PreA = a Then
        Else
            For q = 31 - a To 0 Step -1
                Line1(q).BorderColor = RGB(255, 255, 0)
            Next q
        End If
        PreA = a
    Case Timer3.Interval = 1
        a = Int(Len(Mid) / Zong * 31 + 1)
        If PreA = a Then
        Else
            For q = 31 - a To 0 Step -1
                Line1(q).BorderColor = RGB(0, 0, 255)
            Next q
        End If
        PreA = a
    Case Timer4.Interval = 1
        a = Int(Len(Mid) / Zong * 31 + 1)
        If PreA = a Then
        Else
            For q = 31 - a To 0 Step -1
                Line1(q).BorderColor = RGB(0, 255, 0)
            Next q
        End If
        PreA = a
End Select
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "г��ת��"
   ClientHeight    =   6432
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9672
   BeginProperty Font 
      Name            =   "����"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6432
   ScaleWidth      =   9672
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��"
      Height          =   972
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   492
   End
   Begin VB.CommandButton C 
      Caption         =   "ת��"
      Height          =   732
      Left            =   4170
      MouseIcon       =   "Form1.frx":044A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2760
      Width           =   1332
   End
   Begin VB.TextBox R 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2652
      Left            =   870
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3450
      Width           =   7932
   End
   Begin VB.TextBox T 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2652
      Left            =   870
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   330
      Width           =   7932
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s(2 To 20) As String
Dim p(2 To 20) As String

Private Sub C_Click()
R = ""
Dim i As Integer, q As Byte
Dim b As Boolean
For i = 1 To Len(T)
    For q = 2 To 20
        If InStr(s(q), Mid(T, i, 1)) <> 0 Then
            R = R & p(q)
            b = True
            GoTo TG
        End If
    Next q
TG:    If b = False Then
        R = R & Mid(T, i, 1)
    End If
    b = False
Next i
End Sub

Private Sub Form_Load()
For i = 2 To 9
    p(i) = i
Next i
p(10) = "0"
p(11) = "B"
p(12) = "C"
p(13) = "D"
p(14) = "E"
p(15) = "G"
p(16) = "I"
p(17) = "O"
p(18) = "P"
p(19) = "R"
p(20) = "T"
s(2) = "��������������"
s(3) = "��ɢɡ������"
s(4) = "����˿˺��˽˻˼��˾˹����������ʳ����"
s(5) = "������������������������������������"
s(6) = "����������������������"
s(7) = "�������������������������������������������������������������"
s(8) = "�Ѱ˰ɰְΰհϰͰ԰ǰŰȰ���"
s(9) = "�;žƾɾþ��Ⱦ��˾���"
s(10) = "��������������������������������������������������������"
s(11) = "�ȱʱձǱ̱رܱƱϱ۱˱ɱڱͱұױٱαбӱֱݱѱ��������"
s(12) = "��ϴϸ��ϷϵϲϯϡϪϨ��ϥϢϮϧϰ��Ϧ"
s(13) = "�صڵ׵͵ֵܵεϵݵٵ۵յе޵̵�"
s(14) = "һ������������������ҽ��������������������������������������������������Ҽ"
s(15) = "�������ȼ��������Ǽ��������ļͼ����Ƽ�����ϵ�����ʼ���"
s(16) = "��������������������"
s(17) = "żŻŷźŸ��ŽŹ��ک"
s(18) = "��Ƥ��ƥ������ƨ��ƧƣƦ������ơƩ��ا"
s(19) = "����߹"
s(20) = "�����������������������������"
End Sub

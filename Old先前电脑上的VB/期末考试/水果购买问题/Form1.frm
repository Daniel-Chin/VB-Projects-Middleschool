VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "����ˮ������"
   ClientHeight    =   6420
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11556
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   11556
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6252
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   6372
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ���㣡"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6132
      Left            =   6600
      TabIndex        =   0
      Top             =   120
      Width           =   4812
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.FontSize = 10
Text1 = ""
For i = 1 To 50
    s = i * 3 + (50 - i)                                                                      '���㻨����Ǯ
    If s >= 100 And i > 0 And 50 - i > 0 Then                                                 '��ⷽ���Ƿ�ϸ�
        Text1 = Text1 & "���Թ����㽶" & i & "����ƻ��" & (50 - i) & "����" & Chr(13) & Chr(10) '�������
        j = j + 1                                                                              '�ۼӷ�������
    End If
Next i
Text1 = Text1 & "����" & j & "�ַ�����"                                                        '��ӡ��������
End Sub

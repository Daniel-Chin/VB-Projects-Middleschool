VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "��������"
   ClientHeight    =   5004
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7644
   LinkTopic       =   "Form3"
   ScaleHeight     =   5004
   ScaleWidth      =   7644
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1572
      Left            =   1296
      TabIndex        =   2
      Top             =   2880
      Width           =   5052
   End
   Begin VB.TextBox Text1 
      Height          =   612
      Left            =   2496
      TabIndex        =   1
      Top             =   2040
      Width           =   2652
   End
   Begin VB.Label Label1 
      Caption         =   "�������ǵ�һ��ʹ�ñ��������������롣�����趨֮�󲻿ɸ��ġ�����ʧ���룬����ϵ�������� I_am_qnf@163.com��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   714
      TabIndex        =   0
      Top             =   360
      Width           =   6216
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mima As String
Private Sub Command1_Click()
    If InputBox("���ٴ��������롣", "����ȷ��", "") = Mima Then
        Form1.Visible = True
        XieRu Mima
        MsgBox "ע��ɹ���", vbInformation + vbOKOnly, "��ϲ"
        End
    Else
        MsgBox "�����������벻����", vbCritical, "����"
    End If
End Sub
Private Sub Text1_Change()
Mima = Text1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub
Private Sub XieRu(PW As String)
Open App.Path & "\���������\delper\1.txt" For Output As #2
    Print #2, PW
Close #2
End Sub

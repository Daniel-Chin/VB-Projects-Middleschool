VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�������"
   ClientHeight    =   2628
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "����"
      Size            =   25.8
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2628
   ScaleWidth      =   8280
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ���㡣"
      BeginProperty Font 
         Name            =   "����"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   8052
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
v = 0: p = 0
For i = 1 To 300
    s = i * (600 - 2 * i) / 4        '�������
    If s > v Then                    '�Ƚ�����������̨����ȷ�����ֵ
        v = s
        p = i
    End If
Next i
Print "������ʱ����Ϊ��" & ((600 - 2 * p) / 4) & "cm��" & Chr(13) & "��Ϊ��" & p & "cm��" & Chr(13) & "�����" & v & "cm^2��"
End Sub

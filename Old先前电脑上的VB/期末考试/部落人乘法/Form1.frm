VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�����˳˷�"
   ClientHeight    =   5256
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   9132
   LinkTopic       =   "Form1"
   ScaleHeight     =   5256
   ScaleWidth      =   9132
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "����"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "������"
      Top             =   4440
      Width           =   4812
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "����"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "������"
      Top             =   3240
      Width           =   4812
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���㣡"
      BeginProperty Font 
         Name            =   "����"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   4092
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4920
      TabIndex        =   2
      Text            =   "15"
      Top             =   1440
      Width           =   1452
   End
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
      Height          =   492
      Left            =   1680
      TabIndex        =   1
      Text            =   "13"
      Top             =   1440
      Width           =   1452
   End
   Begin VB.Label Label4 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "����"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   6
      Top             =   4440
      Width           =   252
   End
   Begin VB.Label Label3 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "����"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   5
      Top             =   3240
      Width           =   252
   End
   Begin VB.Label Label2 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3840
      TabIndex        =   3
      Top             =   1440
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "����һ���˷�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8892
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim first_ck As Boolean
first_ck = True                '��һ���ۼ�
Text3 = ""
i = Val(Text1)
p = Val(Text2)
Do While i >= 1                '������>1
    If i Mod 2 = 1 Then        '������Ϊ����
        s = s + p              '�����ۼ��ۼ�
        If first_ck = True Then        '����+��������
            Text3 = Text3 & p          '���ɼӷ���ʽ
        Else
            Text3 = Text3 & "+" & p
        End If
        first_ck = False
    End If
    i = Int(i / 2)             '���м���
    p = 2 * p                  '���мӱ�
Loop
Text4 = s                      '�����
End Sub

Private Sub Form_Load()
Call Command1_Click
End Sub

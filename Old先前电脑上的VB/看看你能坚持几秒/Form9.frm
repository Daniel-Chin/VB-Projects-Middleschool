VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "�浵"
   ClientHeight    =   1956
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8436
   LinkTopic       =   "Form9"
   ScaleHeight     =   1956
   ScaleWidth      =   8436
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "ˢ��"
      Height          =   372
      Left            =   7200
      TabIndex        =   7
      Top             =   240
      Width           =   972
   End
   Begin VB.TextBox Text3 
      Height          =   372
      Left            =   5640
      TabIndex        =   6
      Text            =   "��ѡ��浵��"
      Top             =   1440
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   3000
      TabIndex        =   5
      Text            =   "��ѡ��浵��"
      Top             =   1440
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Text            =   "��ѡ��浵��"
      Top             =   1440
      Width           =   1572
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�浵3"
      Height          =   372
      Left            =   6000
      TabIndex        =   3
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�浵2"
      Height          =   372
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�浵1"
      Height          =   372
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "����Ϸ�ṩ�����浵λ����ѡ��һ���浵��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   6012
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Beep
qr = MsgBox("ȷ��Ҫ�浵��", vbOKCancel, "��ȷ��")
If qr = 1 Then
Else
    Exit Sub
End If
Dim n_m As String
n_m = App.Path & "\saves\1.txt"
Open n_m For Output As #1
Print #1, Text1
Print #1, Form1.money
Print #1, Form1.exp
Print #1, ""       '�������ɾ͡������߷�
Print #1, Form1.x_s
Close
End Sub
Private Sub Command2_Click()
Beep
qr = MsgBox("ȷ��Ҫ�浵��", vbOKCancel, "��ȷ��")
If qr = 1 Then
Else
    Exit Sub
End If
Dim n_m As String
n_m = App.Path & "\saves\2.txt"
Open n_m For Output As #1
Print #1, Text2
Print #1, Form1.money
Print #1, Form1.exp
Print #1, ""       '�������ɾ͡������߷�
Print #1, Form1.x_s
Close
End Sub
Private Sub Command3_Click()
Beep
qr = MsgBox("ȷ��Ҫ�浵��", vbOKCancel, "��ȷ��")
If qr = 1 Then
Else
    Exit Sub
End If
Dim n_m As String
n_m = App.Path & "\saves\3.txt"
Open n_m For Output As #1
Print #1, Text3
Print #1, Form1.money
Print #1, Form1.exp
Print #1, ""       '�������ɾ͡������߷�
Print #1, Form1.x_s
Close
End Sub

Private Sub Command4_Click()
Dim a As String
Dim b As String
Dim c As String
Open App.Path & "\saves\1.txt" For Input As #1
Input #1, a
Close
Open App.Path & "\saves\2.txt" For Input As #2
Input #2, b
Close
Open App.Path & "\saves\3.txt" For Input As #3
Input #3, c
Close
Text1 = a
Text2 = b
Text3 = c
End Sub

Private Sub Form_Load()
Call Command4_Click
End Sub

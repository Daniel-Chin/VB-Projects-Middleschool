VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Բ���ʼ�����"
   ClientHeight    =   6372
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11064
   LinkTopic       =   "Form1"
   ScaleHeight     =   6372
   ScaleWidth      =   11064
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   6240
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Left            =   5880
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Left            =   5520
      Top             =   2040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   372
      Left            =   6120
      TabIndex        =   1
      Top             =   3000
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ��Pie"
      Height          =   372
      Left            =   5040
      TabIndex        =   0
      Top             =   3000
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Բ���ʣ�"
      Height          =   372
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label1 
      Height          =   372
      Left            =   6600
      TabIndex        =   2
      Top             =   600
      Width           =   2772
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Long
Dim s As Double
Dim b As Long
Dim k As Long
Dim c As Long

Private Sub Command1_Click()
c = Val(InputBox("��������ϣ���õ���Բ���ʵľ�ȷ�̶ȡ�����Խ��Խ��ȷ�������벻���� 214748363����Ȼ������Ȼ����Ļ�����û��ϵ��", "��ȷ��", "10000000"))
s = 0
For a = 1 To c
    b = (a * 2 - 1) * (-1) ^ (a + 1)
    s = s + 1 / b
Next a
Print s * 4
End Sub
Private Sub Command2_Click()
Cls
End Sub

Private Sub Command3_Click()
If Timer1.Interval = 0 Then
    Timer1.Interval = 1
Else
    Timer1.Interval = 0
End If
End Sub
Private Sub Timer1_Timer()
Call Command1_Click
End Sub

Private Sub Timer2_Timer()
Cls
End Sub

Private Sub Timer3_Timer()
Label1 = s * 4
End Sub

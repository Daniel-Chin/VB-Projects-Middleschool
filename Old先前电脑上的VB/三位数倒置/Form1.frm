VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1464
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   2280
   LinkTopic       =   "Form1"
   ScaleHeight     =   1464
   ScaleWidth      =   2280
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2052
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   2052
   End
   Begin VB.Label Label1 
      Caption         =   "��text1������һ����λ����������ȷ������"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1932
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Then
    Label1.Caption = "��text1������һ����λ����������ȷ������"
Else
    If Text1 < 1000 And Text1 > 99 Then
        a = Text1 \ 100
        c = Text1 Mod 10
        b = (Text1 Mod 100) \ 10
        Label1.Caption = a + b * 10 + c * 100
    Else
        Call MsgBox("�Բ���������Ĳ�����λ����", vbExclamation, "����")
    End If
End If
End Sub

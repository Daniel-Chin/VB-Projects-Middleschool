VERSION 5.00
Begin VB.Form Starter 
   Caption         =   "����ѧ�����������ɵ�"
   ClientHeight    =   6324
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10644
   LinkTopic       =   "Form1"
   ScaleHeight     =   6324
   ScaleWidth      =   10644
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton YLW 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   6000
      TabIndex        =   2
      Top             =   4440
      Width           =   2412
   End
   Begin VB.CommandButton SMW 
      Caption         =   "˵����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   3960
      TabIndex        =   1
      Top             =   2040
      Width           =   2412
   End
   Begin VB.CommandButton JXW 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2412
   End
End
Attribute VB_Name = "Starter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ZhongXin As String, JiXuDuiXiang As String

Private Sub JXW_Click()
MsgBox "�����£��������²�Ρ��������ɾ䡣Ȧ���ؼ��������Դ�������ҳ����ģ����߹۵������Ʒ�������������ѡ���⣬���ȰѸ���ѡ���һ�飬���ܻ���������", vbOKOnly, "������"
If MsgBox("���������»���־�ˣ�����ѡ��OK��־��ѡ��Cancel��", vbOKCancel + vbQuestion, "������") = vbOK Then
    '����
Else
    '־��
End If
End Sub


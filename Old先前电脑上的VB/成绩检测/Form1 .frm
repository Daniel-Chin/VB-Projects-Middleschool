VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ɼ����"
   ClientHeight    =   3984
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   10788
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   42
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3984
   ScaleWidth      =   10788
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "�鿴�ʵ��б�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   8160
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ָ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5520
      TabIndex        =   4
      Top             =   3204
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8160
      TabIndex        =   2
      Text            =   "�ڴ������ؼ�"
      Top             =   3120
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   25.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Left            =   1680
      TabIndex        =   0
      Text            =   "0"
      Top             =   3120
      Width           =   6372
   End
   Begin VB.Label Label1 
      Caption         =   "�ɼ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim caidan As Double

Private Sub Command1_Click()
MsgBox "�����ǲʵ��б��Ҽ����������Ҫ�˳����ҿ�����ô�����㵰������˭��������ģ���骷��˧�����ꣻ�٣���ɫ����ָ�������ʵ������⣺��һ�ν��뱾����ʱֻҪ���ؼ�����������롰SAGA�����ɡ�", vbOKOnly, "���������"
End Sub

Private Sub text1_Change()
Form1.ForeColor = RGB(0, 0, 0)
a = Val(Text1)
Cls
If a = 0 Then Print "�㵰"
If a > 0 And a < 60 Then Print "������"
If a >= 60 And a < 90 Then Print "�ϸ�"
If a >= 90 And a < 100 Then Print "����"
If a = 100 Then Print "����"
If a > 100 Or a < 0 Then Print "�����������ĳɼ���"
If Text1 = "��Ҫ�˳�" Then MsgBox "����˵Ҫ�˳��ģ�", , "�˳���": End
If Text1 = "�ҿ�����ô�����㵰" Then Cls: Call �ʵ�: Print "���������Լ��������֣�": Call �ʵ�
If Text1 = "�ʵ�" Then Cls: Print "����Ϊ����������": Print "�����вʵ�ô��"
If Text1 = "�������" Then Cls: Call �ʵ�:    Print "�㲻����ô��": Call �ʵ�
If Text1 = "����" Then Cls: Call �ʵ�:    Print "�ҽ�����������������֡���": Call �ʵ�
If Text1 = "��骷��˧" Then Cls:  Call �ʵ�:   Print "лл����Ҳ��˧��": Call �ʵ�
If Text1 = "����" Then Cls: Call �ʵ�:    Print "�������ˣ�": Call �ʵ�
If Text1 = "��" Then Cls: Call �ʵ�:    Print "�����㣬˵�������ã���": Call �ʵ�
If Text1 = "������" Then Cls: Call �ʵ�:    Print "��������": Call �ʵ�
If Text1 = "ָ����" Then Cls: Call �ʵ�:    Print "Speak Friend and enter��": Call �ʵ�
If Text1 = "Friend" Then Cls: Call �ʵ�:    Print "���������������ȷ��": Print "������ֹһ����": Call �ʵ�
If Text1 = "Mellon" Then Cls: Call �ʵ�:    Print "�������Ѿ�ͨ����ָ�������ˣ�": Call �ʵ�: Command2.Visible = True
If Text1 = "����˭" Then Cls:  Call �ʵ�:   Print "�ҵ����ֽ���骷㡣": Call �ʵ�
If Text1 = "�ռ�����" Then Cls: Call �ʵ�:    Print "�л����񹲺͹����꣡����": Call �ʵ�
If Text1 = "�Ҽ���������" Then Cls: Call �ʵ�:    Print "����������ͱ�㰴ť��": Call �ʵ�
If Text1 = "��ɫ����" Then Cls: Call �ʵ�:    Print "������ʴ���": Call �ʵ�
If Text1 = "Four is enough" Then Cls: Call �ʵ�:    Print "ԭ���������˼���ʷ��������": Call �ʵ�
End Sub

Private Sub �ʵ�()
If Text1 = "�ռ�����" Then Form1.ForeColor = RGB(255, 0, 0) Else Form1.ForeColor = RGB(0, 0, 0)
Command1.Visible = True
If caidan = 1 Then
Else
    caidan = 1
    Call MsgBox("��ϲ�㷢�ֲʵ��������ǲʵ��б���Ҫ�˳����ҿ�����ô�����㵰������˭��������ģ���骷��˧�����ꣻ�٣���ɫ�����ʵ���ָ���������⣺��һ�ν��뱾����ʱֻҪ���ؼ�����������롰SAGA�����ɡ�", 0, "�ʵ�")
End If
End Sub

Private Sub Text2_Change()
If Text2 = "SAGA" Then Call �ʵ�
End Sub

Private Sub Command2_Click()
    MsgBox "Come in soon!������һ���ռ����زʵ���", , "��Ҳ��ϲ��ָ������"
End Sub

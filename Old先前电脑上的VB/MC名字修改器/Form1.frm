VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MC�����޸���"
   ClientHeight    =   4128
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7296
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4128
   ScaleWidth      =   7296
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DL As String

Private Sub Form_Load()
Dim S As String
DL = Chr(13) & Chr(10)
On Error GoTo CuoWu
If 0 Then
CuoWu:      MsgBox "��ѱ�������mclauncher.cfg����ͬһ���˵��¡����������ļ�������ǡ���ʼ��Ϸ.exe����", vbCritical + vbOKOnly, "����ʧ��"
            End
End If
Open App.Path & "\mclauncher.cfg" For Input As #1
Dim Ori(1 To 8) As String, yN As String
For i = 1 To 8
    Input #1, Ori(i)
Next i
yN = Right(Ori(2), Len(Ori(2)) - 5)
Dim ShuRu As String
ShuRu = InputBox("������ר��Ӧ��MC���������ù�������Ϸ�������һ���ַ��޷����������⡣������Ҫ�����֡�", "", yN)
Close #1
If ShuRu = "" Then MsgBox "��������û�б��޸ġ�": End
S = Ori(1) & DL & "��Ϸ����=" & ShuRu & DL & Ori(3) & DL & Ori(4) & DL & Ori(5) & DL & Ori(6) & DL & Ori(7) & DL & Ori(8) & DL
On Error Resume Next
Kill App.Path & "\mclauncher.cfg"
Open App.Path & "\mclauncher.cfg" For Output As #1
Print #1, S
Close #1
On Error GoTo ZBD
If MsgBox("�޸ĳɹ�������������Ϸ��", vbYesNo + vbQuestion, "") = vbYes Then
    Shell App.Path & "\��ʼ��Ϸ.exe"
End If
If 0 Then
ZBD:    MsgBox "��ʼ��Ϸ.exe δ�ҵ�������OK���˳� MC�����޸�����ע��������Ϸ���Ѿ��ɹ��޸��ˡ�", vbCritical + vbOKOnly, "����"
End If
End
End Sub

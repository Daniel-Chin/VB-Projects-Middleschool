VERSION 5.00
Begin VB.Form Form3 
   ClientHeight    =   5748
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7644
   LinkTopic       =   "Form3"
   ScaleHeight     =   5748
   ScaleWidth      =   7644
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3360
      Top             =   960
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Byte
Dim Password As String, Logon As Boolean
Private Sub Form_KeyPress(KeyAscii As Integer)
Password = Password & KeyAscii
If Password = "67810" Then
    Form1.Visible = True
    Logon = True
    Unload Form3
End If
End Sub

Private Sub Form_Load()
Show
Me.FontSize = 15
End Sub


Private Sub Timer2_Timer()
t = t + 1
Select Case t
    Case 1
        Print "��ȡSetUp.dll..."
    Case 2
        Print "��ȡLog.dll..."
    Case 3
        Print "������..."
    Case 4
        Print "ERROR!"
    Case 5
        Print "�����𻵣������ܴ򿪡�����3����Զ��رա�"
    Case 6
        Print "2.5������رա�"
    Case 7
        Print "2������رա�"
    Case 8
        Print "1.5������رա�"
    Case 9
        Print "1������رա�"
    Case 10
        Print "0.5������رա�"
    Case 11
        Print "����رա�"
    Case 12
        End
End Select
End Sub

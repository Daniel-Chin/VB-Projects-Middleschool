VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "��Ƨ�ַ�ASCii���ѯ"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3996
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   3996
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1452
      Left            =   732
      TabIndex        =   0
      Top             =   432
      Width           =   2532
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Text1) = 1 Then
        MsgBox QiuMa(Text1), vbOKOnly, "��ѯ���"
    Else
        MsgBox "�����뵥���ַ���", vbCritical, "����"
    End If
End If
End Sub
Public Function QiuMa(Text As String)
Dim i As Long
    For i = -29999 To 29999
        If Chr(i) = Text Then
            QiuMa = i
            Exit For
        End If
    Next i
End Function

VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   ClientHeight    =   0
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   0
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   -240
      Top             =   0
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Password As String, Logon As Boolean
Private Sub Form_KeyPress(KeyAscii As Integer)
Password = Password & KeyAscii
If Password = "67810" Then
    Form1.Visible = True
    Logon = True
    Unload Form3
End If
End Sub

Private Sub Timer1_Timer()
If Logon = False Then
    End
End If
End Sub

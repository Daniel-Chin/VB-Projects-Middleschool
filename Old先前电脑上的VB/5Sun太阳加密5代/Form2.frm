VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   0
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   0
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   -240
      Top             =   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Private Sub Form_KeyPress(KeyAscii As Integer)
A = A & KeyAscii
If A = "67810" Then
    Form1.Visible = True
    Unload Form2
End If
End Sub
Private Sub Form_Load()
MsgBox "Under Cunstruction!", vbCritical + vbOKOnly, "Warning"
End Sub
Private Sub Timer1_Timer()
End
End Sub

VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2316
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3624
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   960
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = Screen.Height
Me.Width = Screen.Width
End Sub
Private Sub Timer1_Timer()
Form2.Visible = True
Unload Form1
ShowCursor False
End Sub


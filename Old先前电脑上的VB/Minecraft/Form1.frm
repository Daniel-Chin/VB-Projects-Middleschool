VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10068
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   10068
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   2040
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   2760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim P(0 To 4) As String, C As Byte, t As Byte

Private Sub Form_Click()
Timer1.Interval = 1
Timer2.Interval = 1
Timer3.Interval = 1
End Sub

Private Sub Form_Load()
With Screen
    Me.Height = .Height
    Me.Width = .Width
End With
Show
P(0) = "¡¡¡ö¡¡¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡¡¡¡¡¡¡ö¡ö¡¡¡¡¡ö¡¡¡¡¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡¡¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡¡¡¡"
P(1) = "¡ö¡¡¡ö¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡¡¡¡¡¡¡¡¡¡¡ö¡¡¡ö¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡ö¡¡¡ö¡¡¡¡¡¡¡ö¡ö¡¡¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡ö¡¡"
P(2) = "¡ö¡ö¡ö¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡¡¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡¡¡¡¡¡¡ö¡ö¡ö¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡ö¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡¡¡¡¡ö¡ö¡ö¡¡¡ö¡¡¡ö¡¡"
P(3) = "¡ö¡¡¡ö¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡¡¡¡¡¡¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡¡¡¡¡¡¡ö¡¡¡ö¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡ö¡¡¡ö¡¡¡¡¡¡¡ö¡ö¡¡¡¡¡ö¡¡¡¡¡¡¡ö¡¡¡ö¡¡"
P(4) = "¡ö¡¡¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡¡¡¡¡¡¡ö¡ö¡¡¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡¡¡ö¡¡¡ö¡ö¡ö¡¡¡ö¡ö¡¡¡¡"
End Sub

Private Sub Timer1_Timer()
t = t + 1
If t = 55 Then
    t = 1
    C = C + 1
    Print ""
End If
If C = 5 Then
    Timer1.Interval = 0
    Timer2.Interval = 0
    Timer3.Interval = 0
    Exit Sub
End If
Print Mid(P(C), t, 1);
End Sub

Private Sub Timer2_Timer()
t = t + 1
If t = 55 Then
    t = 1
    C = C + 1
    Print ""
End If
If C = 5 Then
    Timer1.Interval = 0
    Timer2.Interval = 0
    Timer3.Interval = 0
    Exit Sub
End If
Print Mid(P(C), t, 1);
End Sub

Private Sub Timer3_Timer()
t = t + 1
If t = 55 Then
    t = 1
    C = C + 1
    Print ""
End If
If C = 5 Then
    Timer1.Interval = 0
    Timer2.Interval = 0
    Timer3.Interval = 0
    Exit Sub
End If
Print Mid(P(C), t, 1);
End Sub

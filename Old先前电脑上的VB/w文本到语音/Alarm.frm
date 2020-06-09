VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5724
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9924
   LinkTopic       =   "Form1"
   ScaleHeight     =   5724
   ScaleWidth      =   9924
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   5640
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Height          =   852
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   972
   End
   Begin VB.Timer Timer2 
      Left            =   6960
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   4440
      Top             =   2640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Boolean
Private Sub Form_Resize()
Command1.Top = 0
Command1.Height = Me.Height
End Sub

Private Sub Timer1_Timer()
If Second(Now) Mod 2 = 0 And a Then
    a = False
    Me.BackColor = vbGreen
    Command1.Left = Me.Width \ 10
    Timer2.Interval = 100
End If
If Second(Now) Mod 2 = 1 Then
    a = True
End If
End Sub

Private Sub Timer2_Timer()
Timer2.Interval = 0
Me.BackColor = 0
End Sub

Private Sub Timer3_Timer()
Command1.Left = Command1.Left + (Me.Width \ 40)
End Sub

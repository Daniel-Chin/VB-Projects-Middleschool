VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   1728
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5832
   LinkTopic       =   "Form1"
   ScaleHeight     =   1728
   ScaleWidth      =   5832
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "5"
      Top             =   720
      Width           =   972
   End
   Begin VB.Timer Timer1 
      Left            =   4080
      Top             =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   372
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click me!"
      Height          =   372
      Left            =   2040
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "Ê£ÓàÉúÃü£º"
      Height          =   252
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double
Private Sub Command1_Click()
Text1 = Val(Text1) + 1
End Sub
Private Sub Command2_Click()
Timer1.Interval = 10
a = -30
Command1.Visible = True
Command1.Left = Rnd() * (Form1.Width - 173)
Text1 = Text1 - 1
End Sub
Private Sub Timer1_Timer()
If Text1 < 0 Then MsgBox ("ÄãËÀÁË"): End
If a < 30 Then
    a = a + 1
    Command1.Top = -a ^ 2 + 400
Else
    Timer1.Interval = 0
    a = -20
    Command1.Visible = False
End If
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   10020
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6240
      Top             =   4800
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   3960
      Top             =   4680
   End
   Begin VB.Label WQD 
      Caption         =   "¡Þ"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   25.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   492
      Left            =   4164
      TabIndex        =   2
      Top             =   2800
      Width           =   516
   End
   Begin VB.Label ZBZ 
      Height          =   372
      Left            =   3720
      TabIndex        =   1
      Top             =   2880
      Width           =   492
   End
   Begin VB.Label ZS 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   25.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6360
      TabIndex        =   0
      Top             =   1080
      Width           =   2892
   End
   Begin VB.Image Image1 
      Height          =   3672
      Left            =   840
      Picture         =   "Form1.frx":0000
      Top             =   480
      Width           =   7092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer2.Interval = 400
End Sub

Private Sub Timer1_Timer()
ZBZ.Left = ZBZ.Left + 5
ZBZ.Width = ZBZ.Width - 5
On Error GoTo ass
WQD.ForeColor = WQD.ForeColor + &H10101 * 3
If 0 Then
ass:
Timer1 = False
Timer2 = False
End If
End Sub

Private Sub Timer2_Timer()
Timer2.Interval = Timer2.Interval * 0.9
ZS = Val(ZS) + 1
If Timer2.Interval <= 100 Then ZS = ZS * 2
End Sub

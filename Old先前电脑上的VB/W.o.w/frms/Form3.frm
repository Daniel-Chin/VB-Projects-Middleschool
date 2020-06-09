VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000008&
   Caption         =   "信息确认"
   ClientHeight    =   3468
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   7068
   LinkTopic       =   "Form3"
   ScaleHeight     =   3468
   ScaleWidth      =   7068
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   5880
      Top             =   1320
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   372
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   2412
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   492
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Width           =   2412
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   $"Form3.frx":0000
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   2412
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   $"Form3.frx":0016
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3372
   End
   Begin VB.Image Image2 
      Height          =   1896
      Left            =   840
      Picture         =   "Form3.frx":004E
      Top             =   1440
      Width           =   4944
   End
   Begin VB.Image Image1 
      Height          =   408
      Left            =   5880
      Picture         =   "Form3.frx":655C
      Top             =   0
      Width           =   1188
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public player As String
Public playercd As String
Dim na As String
Dim nmal As Boolean
Private Sub Form_Unload(Cancel As Integer)
If nmal = True Then
Else
    End
End If
End Sub
Private Sub Form_Load()
nmal = False
End Sub

Private Sub Image2_Click()
player = Text1
playercd = Text2
nmal = True
Unload Form3
Form4.Visible = True
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        na = Text1
        Text2.SetFocus
        Timer1.Interval = 1
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Image2_Click
    End If
End Sub

Private Sub Timer1_Timer()
Text1 = na
Timer1.Interval = 0
End Sub

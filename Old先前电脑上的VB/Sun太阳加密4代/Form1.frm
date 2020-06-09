VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Ì«Ñô¼ÓÃÜ"
   ClientHeight    =   7104
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12768
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7104
   ScaleWidth      =   12768
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2412
      Left            =   1218
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":74F2
      Top             =   960
      Width           =   10332
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Index           =   3
      X1              =   840
      X2              =   11520
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Index           =   2
      X1              =   11760
      X2              =   11760
      Y1              =   600
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Index           =   1
      X1              =   12240
      X2              =   1680
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Index           =   0
      X1              =   960
      X2              =   960
      Y1              =   1200
      Y2              =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.BackColor = RGB(0, 0, 50)
With Text1
    .BackColor = RGB(0, 0, 50)
    .SelStart = 0
    .SelLength = 5
End With
End Sub

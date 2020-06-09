VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   4308
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5472
   LinkTopic       =   "Form1"
   ScaleHeight     =   4308
   ScaleWidth      =   5472
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4680
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   4680
      Top             =   600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Left            =   2280
      TabIndex        =   0
      Top             =   1800
      Width           =   852
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   2640
      X2              =   2640
      Y1              =   480
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   1080
      X2              =   4200
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderWidth     =   10
      FillColor       =   &H00C0C0C0&
      Height          =   3492
      Left            =   720
      Shape           =   3  'Circle
      Top             =   360
      Width           =   3852
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   25
      FillColor       =   &H00C0C0C0&
      Height          =   3492
      Left            =   720
      Shape           =   3  'Circle
      Top             =   360
      Width           =   3852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer, c As Integer, d As Long, e As Long, f As Long, g As Long


Private Sub Form_Load()
a = 6
c = 500
d = Label1.Left
e = Label1.Top
f = Shape1.Left
g = Shape1.Top
End Sub

Private Sub Timer1_Timer()
If Rnd() > 0.8 Then
    b = Int(Rnd * 100) + 155
    Line1.BorderColor = RGB(b, b, b)
    Line2.BorderColor = Line1.BorderColor
End If
If Rnd() > 0.8 Then
    b = Int(Rnd * 100)
    Form1.BackColor = RGB(b, b, b)
End If
If Rnd() > 0.8 Then
    b = Int(Rnd * 100) + 50
    Shape2.BorderColor = RGB(b, b, b)
End If
If Rnd() > 0.8 Then
    b = Int(Rnd * 100) + 100
    Shape1.BorderColor = RGB(b, b, b)
End If
If Rnd() > 0.9 Then
    b = Int(Rnd * c)
    Label1.Left = Label1.Left + b - c / 2
    b = Int(Rnd * c)
    Label1.Top = Label1.Top + b - c / 2
End If
If Rnd() > 0.9 Then
    b = Int(Rnd * c)
    Shape1.Left = Shape1.Left + b - c / 2
    b = Int(Rnd * c)
    Shape1.Top = Shape1.Top + b - c / 2
End If
End Sub

Private Sub Timer2_Timer()
a = a - 1
Label1.Caption = a
Label1.Left = d
Label1.Top = e
Shape1.Left = f
Shape1.Top = g
End Sub

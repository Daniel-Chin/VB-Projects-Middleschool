VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8772
   ClientLeft      =   132
   ClientTop       =   480
   ClientWidth     =   16296
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8772
   ScaleWidth      =   16296
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4560
      Top             =   3000
   End
   Begin VB.Image Img 
      Height          =   2292
      Left            =   720
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4812
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const CD As Double = 0.1404494382

Public XXX, YYY

Private Sub Form_Load()
Show
XXX = 100000
YYY = 60000
Form2.Show
Form2.Move Screen.Width * 0.8, 0, XXX, YYY
Me.Move 0, 0, XXX, YYY
Me.AutoRedraw = True
Me.FontSize = 1000
Me.Print "diva"
Me.ForeColor = vbWhite
End Sub

Private Sub Timer1_Timer()
Dim X, Y, XX, YY
Do Until IsOn(X, Y)
    X = Int(Rnd() * XXX)
    Y = Int(Rnd() * YYY)
    'DoEvents
Loop
Do Until IsOn(XX, YY)
    XX = X + Int((Rnd() - 0.5) * CD * XXX)
    YY = Y + Int((Rnd() - 0.5) * CD * XXX)
    'DoEvents
Loop
Form2.Line (X, Y)-(XX, YY)
Img.Picture = Form2.Image
DoEvents
End Sub

Function IsOn(X, Y) As Boolean
IsOn = Point(X, Y) = vbWhite
End Function

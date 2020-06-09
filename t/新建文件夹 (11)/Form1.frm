VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   10692
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   10692
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5520
      Top             =   4440
   End
   Begin VB.Image Image1 
      Height          =   2412
      Left            =   600
      Stretch         =   -1  'True
      Top             =   720
      Width           =   7452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Root As String = "D:\VB\How 2 Become A Calculater\Install\"
Dim Tms As Single
Dim OShade As Integer
Const Cy As Integer = 4-

Private Sub Form_Load()
Timer1.Interval = 40
End Sub

Private Sub Timer1_Timer()
Tms = Tms + Timer1.Interval / 1000
If Tms >= Cy Then
    Tms = Tms - Cy
End If
Dim Shade As Integer
Shade = Int((Cos(Tms / Cy * 2 * 3.14159265358979) + 1) / 2 * 10)
If Shade = 10 Then Shade = 9
If Shade <> OShade Then
    Me.Caption = Rnd
    Me.PaintPicture LoadPicture(Root & Shade & ".jpg"), 0, 0
    OShade = Shade
End If
End Sub

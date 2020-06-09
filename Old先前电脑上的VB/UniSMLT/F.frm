VERSION 5.00
Begin VB.Form F 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Universe Simulator"
   ClientHeight    =   6264
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   6264
   ScaleWidth      =   10320
   Begin VB.Timer Main 
      Left            =   9120
      Top             =   3960
   End
   Begin VB.Shape C 
      BackColor       =   &H00000000&
      BorderColor     =   &H0000C000&
      BorderWidth     =   3
      FillColor       =   &H0000FF00&
      FillStyle       =   3  'Vertical Line
      Height          =   1092
      Index           =   2
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   1452
   End
   Begin VB.Shape C 
      BackColor       =   &H00000000&
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      FillColor       =   &H0000FFFF&
      FillStyle       =   2  'Horizontal Line
      Height          =   1092
      Index           =   1
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   1452
   End
   Begin VB.Shape C 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FFFF00&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1092
      Index           =   3
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   1452
   End
End
Attribute VB_Name = "F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Integer
Private Type Star
    Position(0 To 2) As Long
    Speed(0 To 2) As Integer
    Mass As Integer
End Type
Private Type Force
    X As Integer
    Y As Integer
    Z As Integer
    F As Integer
    Angle(0 To 2) As Double
End Type
Dim Star(1 To 3) As Star
Const G As Integer = 50

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 27     'ESC
        End
    Case 32     'SPACE
        Main.Interval = 25 - Main.Interval
End Select
End Sub

Private Sub Form_Load()
With Screen
    Me.Height = .Height
    Me.Width = .Width
End With

'''''''''''''''''''''''
With Star(1)
    .Position(1) = 1369
    .Position(2) = 720
    .Position(0) = 12
    .Speed(1) = 50
    .Speed(2) = -25
    .Mass = 50
End With
With Star(2)
    .Position(1) = 8369
    .Position(2) = 720
    .Position(0) = 30
    .Speed(1) = -20
    .Speed(2) = 25
    .Mass = 30
End With
With Star(3)
    .Position(1) = 2069
    .Position(2) = 2000
    .Position(0) = -95
    .Speed(1) = 12
    .Speed(2) = 18
    .Mass = 5
End With
'''''''''''''''''''''''

Call Main_Timer
End Sub

Private Sub Main_Timer()
Dim F(1 To 3) As Force
For i = 1 To 3  '决定主星
    With Star(i)
        For q = 1 To 3  '决定和谁作用
            If q = i Then
                GoTo TG
            End If
            For p = 0 To 2  '哪个方向上的力
                F(q).F = .Mass * Star(q).Mass / (Dis(Star(i), Star(q))) ^ 2
                F(q).Angle(p) = (Star(q).Position(p) - .Position(p)) / Dis(Star(q), Star(i))
            Next p
            F(q).X = F(q).F * F(q).Angle(1)
            F(q).Y = F(q).F * F(q).Angle(2)
            F(q).Z = F(q).F * F(q).Angle(0)
            For p = 0 To 2
                Add .Speed(p), (F(q).X + F(q).Y + F(q).Z) / .Mass * G
                Add .Position(p), .Speed(p)
            Next p
TG:     Next q
        C(i).Top = .Position(1)
        C(i).Left = .Position(2)
    End With
Next i
End Sub

Private Sub Add(Number1, Number2)
Number1 = Int(Number1 + Number2)
End Sub

Private Function Dis(A As Star, B As Star) As Long
Dis = Sqr((A.Position(0) - B.Position(0)) ^ 2 + (A.Position(1) - B.Position(1)) ^ 2 + (A.Position(2) - B.Position(2)) ^ 2)
End Function

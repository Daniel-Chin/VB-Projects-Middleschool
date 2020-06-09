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
      Left            =   8640
      Top             =   3480
   End
   Begin VB.Timer S3 
      Left            =   6720
      Top             =   4800
   End
   Begin VB.Timer S2 
      Left            =   6000
      Top             =   3240
   End
   Begin VB.Timer S1 
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
Dim asd As Integer
Dim wsx As Integer
Private Type Star
    X As Long
    Y As Long
    Sx As Integer
    Sy As Integer
    mass As Integer
End Type
Dim Star(1 To 3) As Star
Dim CurX As Long, CurY As Long

Private Sub Form_Click()
For i = 1 To 3
    Star(i).X = Int((Star(i).X + CurX) / 2)
    Star(i).Y = Int((Star(i).Y + CurY) / 2)
    Star(i).Sx = 0
    Star(i).Sy = 0
Next i
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 27     'ESC
        End
    Case 32     'SPACE
        For i = 1 To 3
            C(i).Width = Int(Sqr(Star(i).mass) * 100) / 2
            C(i).Height = C(i).Width
        Next i
        Main.Interval = 5 - Main.Interval
        S1.Interval = 5 - S1.Interval
        S2.Interval = 5 - S2.Interval
        S3.Interval = 5 - S3.Interval
End Select
End Sub

Private Sub Form_Load()
With Screen
    Me.Height = .Height
    Me.Width = .Width
End With

'''''''''''''''''''''''
With Star(1)
    .mass = 1
    .X = Screen.Width / 2
    .Y = 500
    .Sx = 150
End With
With Star(2)
    .X = Screen.Width / 2
    .Y = Screen.Height / 2
    .Sx = 30
    .Sy = -32
    .mass = 200
End With
With Star(3)
    .mass = 100
    .X = Star(2).X - 2000
    .Y = Star(2).Y - 2000
    .Sx = -60
    .Sy = 58
End With
'''''''''''''''''''''''

Call Main_Timer
End Sub

Private Function Dis(A As Star, B As Star) As Long
Dis = Sqr((A.X - B.X) ^ 2 + (A.Y - B.Y) ^ 2)
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CurX = X
CurY = Y
End Sub

Private Sub Main_Timer()
For i = 1 To 3
    'C(i).Left = Int(Star(i).X / 2) + 4000
    'C(i).Top = Int(Star(i).Y / 2) + 3000
    C(i).Left = Star(i).X
    C(i).Top = Star(i).Y
Next i
End Sub

Private Sub S1_Timer()
With Star(1)
    p2 = .mass * Star(2).mass / Dis(Star(1), Star(2))
    p2 = p2 * 50
    X2 = p2 * (Star(2).X - .X) / Dis(Star(1), Star(2))
    Y2 = p2 * (Star(2).Y - .Y) / Dis(Star(1), Star(2))
    p3 = .mass * Star(3).mass / Dis(Star(1), Star(3))
    p3 = p3 * 50
    X3 = p3 * (Star(3).X - .X) / Dis(Star(1), Star(3))
    Y3 = p3 * (Star(3).Y - .Y) / Dis(Star(1), Star(3))
    .Sx = .Sx + (X2 + X3) / .mass
    .Sy = .Sy + (Y2 + Y3) / .mass
    .X = .X + .Sx
    .Y = .Y + .Sy
    ''''''''''
    If .X > Screen.Width Then .X = 0
    If .Y > Screen.Width Then .Y = 0
    If .X < 0 Then .X = Screen.Width
    If .Y < 0 Then .Y = Screen.Width
End With
End Sub

Private Sub S2_Timer()
With Star(2)
    p1 = .mass * Star(1).mass / Dis(Star(2), Star(1))
    p1 = p1 * 5
    X1 = p1 * (Star(1).X - .X) / Dis(Star(2), Star(1))
    Y1 = p1 * (Star(1).Y - .Y) / Dis(Star(2), Star(1))
'
    p3 = .mass * Star(3).mass / Dis(Star(2), Star(3))
    p3 = p3 * 50
    X3 = p3 * (Star(3).X - .X) / Dis(Star(2), Star(3))
    Y3 = p3 * (Star(3).Y - .Y) / Dis(Star(2), Star(3))
'
    .Sx = .Sx + (X1 + X3) / .mass
    .Sy = .Sy + (Y1 + Y3) / .mass
'    .X = .X + .Sx
'    .Y = .Y + .Sy
    .X = CurX
    .Y = CurY
    ''''''''''
    If .X > Screen.Width Then .X = 0
    If .Y > Screen.Width Then .Y = 0
    If .X < 0 Then .X = Screen.Width
    If .Y < 0 Then .Y = Screen.Width
End With
End Sub

Private Sub S3_Timer()
With Star(3)
    p2 = .mass * Star(2).mass / Dis(Star(3), Star(2))
    p2 = p2 * 50
    X2 = p2 * (Star(2).X - .X) / Dis(Star(3), Star(2))
    Y2 = p2 * (Star(2).Y - .Y) / Dis(Star(3), Star(2))
'
    p1 = .mass * Star(3).mass / Dis(Star(3), Star(1))
    p1 = p1 * 5
    X1 = p1 * (Star(1).X - .X) / Dis(Star(3), Star(1))
    Y1 = p1 * (Star(1).Y - .Y) / Dis(Star(3), Star(1))
'
    .Sx = .Sx + (X2 + X1) / .mass
    .Sy = .Sy + (Y2 + Y1) / .mass
    .X = .X + .Sx
    .Y = .Y + .Sy
    ''''''''''
    If .X > Screen.Width Then .X = 0
    If .Y > Screen.Width Then .Y = 0
    If .X < 0 Then .X = Screen.Width
    If .Y < 0 Then .Y = Screen.Width
End With
End Sub


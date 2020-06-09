VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12816
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   12816
   Begin VB.Timer BianSe 
      Interval        =   50
      Left            =   6600
      Top             =   3960
   End
   Begin VB.Timer HuanChongQi 
      Left            =   7320
      Top             =   3960
   End
   Begin VB.Timer Main 
      Left            =   6960
      Top             =   3960
   End
   Begin VB.Timer Timer4 
      Left            =   3480
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Left            =   2640
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Left            =   1560
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   480
   End
   Begin VB.Shape Player 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   10
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   612
      Left            =   6840
      Top             =   3840
      Width           =   612
   End
   Begin VB.Shape Enemy 
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   972
      Index           =   3
      Left            =   4080
      Top             =   5880
      Width           =   612
   End
   Begin VB.Shape Enemy 
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   732
      Index           =   2
      Left            =   10320
      Top             =   1680
      Width           =   972
   End
   Begin VB.Shape Enemy 
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   492
      Index           =   1
      Left            =   9720
      Top             =   6360
      Width           =   2412
   End
   Begin VB.Shape Enemy 
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   612
      Index           =   0
      Left            =   3600
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      Visible         =   0   'False
      X1              =   3960
      X2              =   4200
      Y1              =   960
      Y2              =   720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      Visible         =   0   'False
      X1              =   3120
      X2              =   3360
      Y1              =   960
      Y2              =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      Visible         =   0   'False
      X1              =   2160
      X2              =   2400
      Y1              =   960
      Y2              =   720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      Visible         =   0   'False
      X1              =   3960
      X2              =   4200
      Y1              =   840
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      Visible         =   0   'False
      X1              =   3120
      X2              =   3360
      Y1              =   840
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      Visible         =   0   'False
      X1              =   2160
      X2              =   2400
      Y1              =   840
      Y2              =   600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   3960
      X2              =   4200
      Y1              =   720
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   3120
      X2              =   3360
      Y1              =   720
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   2160
      X2              =   2400
      Y1              =   720
      Y2              =   480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   3960
      X2              =   4200
      Y1              =   600
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   3120
      X2              =   3360
      Y1              =   600
      Y2              =   360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   2160
      X2              =   2400
      Y1              =   600
      Y2              =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   3960
      X2              =   4200
      Y1              =   480
      Y2              =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   3120
      X2              =   3360
      Y1              =   480
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   2160
      X2              =   2400
      Y1              =   480
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      Visible         =   0   'False
      X1              =   960
      X2              =   1200
      Y1              =   720
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      Visible         =   0   'False
      X1              =   960
      X2              =   1200
      Y1              =   600
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   960
      X2              =   1200
      Y1              =   480
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   1200
      X2              =   960
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   960
      X2              =   1200
      Y1              =   840
      Y2              =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EA(0 To 3) As Integer
Dim ES As Integer
Const BsXs As Byte = 20
Dim MainR As Byte, MainG As Byte, MainB As Byte, YsCase As Byte
Dim HuanChong As Boolean
Dim GameOn As Boolean
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Dim XX As Integer, YY As Integer
Dim Tms(1 To 4) As Byte
Dim Pi As Double
Dim XiShu As Integer
Dim JiaoDu1(0 To 4) As Integer
Dim JiaoDu2(0 To 4) As Integer
Dim JiaoDu3(0 To 4) As Integer
Dim JiaoDu4(0 To 4) As Integer
Private Sub HuoHua(Which As Byte, Left As Integer, Top As Integer)
Dim SuiJi(0 To 4) As Integer
If Left = 0 Then
    For i = 0 To 4
        SuiJi(i) = Int(Rnd() * 180)
        If SuiJi(i) < 0 Then SuiJi(i) = 360 + SuiJi(i)
    Next i
End If
If Top = 0 Then
    For i = 0 To 4
        SuiJi(i) = Int(Rnd() * 180 + 75)
        If SuiJi(i) < 0 Then SuiJi(i) = 360 + SuiJi(i)
    Next i
End If
If Left > Screen.Width - 100 Then
    For i = 0 To 4
        SuiJi(i) = Int(Rnd() * 180 + 180)
        If SuiJi(i) < 0 Then SuiJi(i) = 360 + SuiJi(i)
    Next i
End If
If Top > Screen.Height - 100 Then
    For i = 0 To 4
        SuiJi(i) = Int(Rnd() * 180 - 85)
        If SuiJi(i) < 0 Then SuiJi(i) = 360 + SuiJi(i)
    Next i
End If
Select Case Which
    Case 1
        For i = 0 To 4
            JiaoDu1(i) = SuiJi(i)
            Line1(i).X1 = Left
            Line1(i).Y1 = Top
            Line1(i).X2 = Left
            Line1(i).Y2 = Top
            Line1(i).Visible = True
        Next i
        Timer1.Interval = 20
    Case 2
        For i = 0 To 4
            JiaoDu2(i) = SuiJi(i)
            Line2(i).X1 = Left
            Line2(i).Y1 = Top
            Line2(i).X2 = Left
            Line2(i).Y2 = Top
            Line2(i).Visible = True
        Next i
        Timer2.Interval = 20
    Case 3
        For i = 0 To 4
            JiaoDu3(i) = SuiJi(i)
            Line3(i).X1 = Left
            Line3(i).Y1 = Top
            Line3(i).X2 = Left
            Line3(i).Y2 = Top
            Line3(i).Visible = True
        Next i
        Timer3.Interval = 20
    Case 4
        For i = 0 To 4
            JiaoDu4(i) = SuiJi(i)
            Line4(i).X1 = Left
            Line4(i).Y1 = Top
            Line4(i).X2 = Left
            Line4(i).Y2 = Top
            Line4(i).Visible = True
        Next i
        Timer4.Interval = 20
End Select
End Sub
Private Sub BianSe_Timer()
Select Case YsCase
    Case 0
        MainR = MainR + BsXs
        If MainR = 240 Then YsCase = 1 'R
    Case 1
        MainG = MainG + BsXs
        If MainG = 240 Then YsCase = 2 'RG
    Case 2
        MainR = MainR - BsXs
        If MainR = 0 Then YsCase = 3 'G
    Case 3
        MainB = MainB + BsXs
        If MainB = 240 Then YsCase = 4 'GB
    Case 4
        MainG = MainG - BsXs
        If MainG = 0 Then YsCase = 5 'B
    Case 5
        MainR = MainR + BsXs
        If MainR = 240 Then YsCase = 6 'BR
    Case 6
        MainB = MainB - BsXs
        If MainB = 0 Then YsCase = 1 'R
End Select
Player.FillColor = RGB(MainR, MainG, MainB)
Player.BorderColor = RGB(255 - MainR, 255 - MainG, 255 - MainB)
End Sub
Private Sub Form_Click()
ShowCursor False
Main.Interval = 20
GameOn = True
End Sub
Private Sub Form_Load()
Randomize
ES = 150
EA(0) = 45
EA(1) = 315
EA(2) = 225
EA(3) = 135
XiShu = 300
Pi = 3.14159265358979
Me.BackColor = RGB(0, 0, 40)
With Screen
    Me.Width = .Width
    Me.Height = .Height
End With
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If GameOn = True And HuanChong = False Then
    Player.Top = Y - Player.Height / 2
    Player.Left = X - Player.Width / 2
    HuanChong = True
    HuanChongQi.Interval = 10
End If
End Sub
Private Sub HuanChongQi_Timer()
HuanChong = False
End Sub
Private Sub Timer1_Timer()
Tms(1) = Tms(1) + 1
For q = 0 To 4
    Select Case JiaoDu1(q)
        Case Is <= 90
            Line1(q).X1 = Line1(q).X1 + XiShu * Sin(JiaoDu1(q) * Pi / 180) / 1.5
            Line1(q).Y1 = Line1(q).Y1 + XiShu * Cos(JiaoDu1(q) * Pi / 180) / 1.5
            Line1(q).X2 = Line1(q).X2 + XiShu * Sin(JiaoDu1(q) * Pi / 180)
            Line1(q).Y2 = Line1(q).Y2 + XiShu * Cos(JiaoDu1(q) * Pi / 180)
        Case Is <= 180 And JiaoDu1(q) > 90
            Line1(q).X1 = Line1(q).X1 + XiShu * Sin((180 - JiaoDu1(q)) * Pi / 180) / 1.5
            Line1(q).Y1 = Line1(q).Y1 - XiShu * Cos((180 - JiaoDu1(q)) * Pi / 180) / 1.5
            Line1(q).X2 = Line1(q).X2 + XiShu * Sin((180 - JiaoDu1(q)) * Pi / 180)
            Line1(q).Y2 = Line1(q).Y2 - XiShu * Cos((180 - JiaoDu1(q)) * Pi / 180)
        Case Is <= 270 And JiaoDu1(q) > 180
            Line1(q).X1 = Line1(q).X1 - XiShu * Cos((270 - JiaoDu1(q)) * Pi / 180) / 1.5
            Line1(q).Y1 = Line1(q).Y1 + XiShu * Sin((270 - JiaoDu1(q)) * Pi / 180) / 1.5
            Line1(q).X2 = Line1(q).X2 - XiShu * Cos((270 - JiaoDu1(q)) * Pi / 180)
            Line1(q).Y2 = Line1(q).Y2 + XiShu * Sin((270 - JiaoDu1(q)) * Pi / 180)
        Case Is > 270
            Line1(q).X1 = Line1(q).X1 - XiShu * Sin((360 - JiaoDu1(q)) * Pi / 180) / 1.5
            Line1(q).Y1 = Line1(q).Y1 - XiShu * Cos((360 - JiaoDu1(q)) * Pi / 180) / 1.5
            Line1(q).X2 = Line1(q).X2 - XiShu * Sin((360 - JiaoDu1(q)) * Pi / 180)
            Line1(q).Y2 = Line1(q).Y2 - XiShu * Cos((360 - JiaoDu1(q)) * Pi / 180)
    End Select
Next q
If Tms(1) = 4 Then
    Tms(1) = 0
    Timer1.Interval = 0
    For i = 0 To 4
        Line1(i).Visible = False
    Next i
End If
End Sub
Private Sub Timer2_Timer()
Tms(2) = Tms(2) + 1
For q = 0 To 4
    Select Case JiaoDu2(q)
        Case Is <= 90
            Line2(q).X1 = Line2(q).X1 + XiShu * Sin(JiaoDu2(q) * Pi / 180) / 2
            Line2(q).Y1 = Line2(q).Y1 + XiShu * Cos(JiaoDu2(q) * Pi / 180) / 2
            Line2(q).X2 = Line2(q).X2 + XiShu * Sin(JiaoDu2(q) * Pi / 180)
            Line2(q).Y2 = Line2(q).Y2 + XiShu * Cos(JiaoDu2(q) * Pi / 180)
        Case Is <= 180 And JiaoDu2(q) > 90
            Line2(q).X1 = Line2(q).X1 + XiShu * Sin((180 - JiaoDu2(q)) * Pi / 180) / 3
            Line2(q).Y1 = Line2(q).Y1 - XiShu * Cos((180 - JiaoDu2(q)) * Pi / 180) / 3
            Line2(q).X2 = Line2(q).X2 + XiShu * Sin((180 - JiaoDu2(q)) * Pi / 180)
            Line2(q).Y2 = Line2(q).Y2 - XiShu * Cos((180 - JiaoDu2(q)) * Pi / 180)
        Case Is <= 270 And JiaoDu2(q) > 180
            Line2(q).X1 = Line2(q).X1 - XiShu * Cos((270 - JiaoDu2(q)) * Pi / 180) / 3
            Line2(q).Y1 = Line2(q).Y1 + XiShu * Sin((270 - JiaoDu2(q)) * Pi / 180) / 3
            Line2(q).X2 = Line2(q).X2 - XiShu * Cos((270 - JiaoDu2(q)) * Pi / 180)
            Line2(q).Y2 = Line2(q).Y2 + XiShu * Sin((270 - JiaoDu2(q)) * Pi / 180)
        Case Is > 270
            Line2(q).X1 = Line2(q).X1 - XiShu * Sin((360 - JiaoDu2(q)) * Pi / 180) / 3
            Line2(q).Y1 = Line2(q).Y1 - XiShu * Cos((360 - JiaoDu2(q)) * Pi / 180) / 3
            Line2(q).X2 = Line2(q).X2 - XiShu * Sin((360 - JiaoDu2(q)) * Pi / 180)
            Line2(q).Y2 = Line2(q).Y2 - XiShu * Cos((360 - JiaoDu2(q)) * Pi / 180)
    End Select
Next q
If Tms(2) = 4 Then
    Tms(2) = 0
    Timer2.Interval = 0
    For i = 0 To 4
        Line2(i).Visible = False
    Next i
End If
End Sub
Private Sub Timer3_Timer()
Tms(3) = Tms(3) + 1
For q = 0 To 4
    Select Case JiaoDu3(q)
        Case Is <= 90
            Line3(q).X1 = Line3(q).X1 + XiShu * Sin(JiaoDu3(q) * Pi / 180) / 2
            Line3(q).Y1 = Line3(q).Y1 + XiShu * Cos(JiaoDu3(q) * Pi / 180) / 2
            Line3(q).X2 = Line3(q).X2 + XiShu * Sin(JiaoDu3(q) * Pi / 180)
            Line3(q).Y2 = Line3(q).Y2 + XiShu * Cos(JiaoDu3(q) * Pi / 180)
        Case Is <= 180 And JiaoDu3(q) > 90
            Line3(q).X1 = Line3(q).X1 + XiShu * Sin((180 - JiaoDu3(q)) * Pi / 180) / 3
            Line3(q).Y1 = Line3(q).Y1 - XiShu * Cos((180 - JiaoDu3(q)) * Pi / 180) / 3
            Line3(q).X2 = Line3(q).X2 + XiShu * Sin((180 - JiaoDu3(q)) * Pi / 180)
            Line3(q).Y2 = Line3(q).Y2 - XiShu * Cos((180 - JiaoDu3(q)) * Pi / 180)
        Case Is <= 270 And JiaoDu3(q) > 180
            Line3(q).X1 = Line3(q).X1 - XiShu * Cos((270 - JiaoDu3(q)) * Pi / 180) / 3
            Line3(q).Y1 = Line3(q).Y1 + XiShu * Sin((270 - JiaoDu3(q)) * Pi / 180) / 3
            Line3(q).X2 = Line3(q).X2 - XiShu * Cos((270 - JiaoDu3(q)) * Pi / 180)
            Line3(q).Y2 = Line3(q).Y2 + XiShu * Sin((270 - JiaoDu3(q)) * Pi / 180)
        Case Is > 270
            Line3(q).X1 = Line3(q).X1 - XiShu * Sin((360 - JiaoDu3(q)) * Pi / 180) / 3
            Line3(q).Y1 = Line3(q).Y1 - XiShu * Cos((360 - JiaoDu3(q)) * Pi / 180) / 3
            Line3(q).X2 = Line3(q).X2 - XiShu * Sin((360 - JiaoDu3(q)) * Pi / 180)
            Line3(q).Y2 = Line3(q).Y2 - XiShu * Cos((360 - JiaoDu3(q)) * Pi / 180)
    End Select
Next q
If Tms(3) = 4 Then
    Tms(3) = 0
    Timer3.Interval = 0
    For i = 0 To 4
        Line3(i).Visible = False
    Next i
End If
End Sub
Private Sub Timer4_Timer()
Tms(4) = Tms(4) + 1
For q = 0 To 4
    Select Case JiaoDu4(q)
        Case Is <= 90
            Line4(q).X1 = Line4(q).X1 + XiShu * Sin(JiaoDu4(q) * Pi / 180) / 2
            Line4(q).Y1 = Line4(q).Y1 + XiShu * Cos(JiaoDu4(q) * Pi / 180) / 2
            Line4(q).X2 = Line4(q).X2 + XiShu * Sin(JiaoDu4(q) * Pi / 180)
            Line4(q).Y2 = Line4(q).Y2 + XiShu * Cos(JiaoDu4(q) * Pi / 180)
        Case Is <= 180 And JiaoDu4(q) > 90
            Line4(q).X1 = Line4(q).X1 + XiShu * Sin((180 - JiaoDu4(q)) * Pi / 180) / 3
            Line4(q).Y1 = Line4(q).Y1 - XiShu * Cos((180 - JiaoDu4(q)) * Pi / 180) / 3
            Line4(q).X2 = Line4(q).X2 + XiShu * Sin((180 - JiaoDu4(q)) * Pi / 180)
            Line4(q).Y2 = Line4(q).Y2 - XiShu * Cos((180 - JiaoDu4(q)) * Pi / 180)
        Case Is <= 270 And JiaoDu4(q) > 180
            Line4(q).X1 = Line4(q).X1 - XiShu * Cos((270 - JiaoDu4(q)) * Pi / 180) / 3
            Line4(q).Y1 = Line4(q).Y1 + XiShu * Sin((270 - JiaoDu4(q)) * Pi / 180) / 3
            Line4(q).X2 = Line4(q).X2 - XiShu * Cos((270 - JiaoDu4(q)) * Pi / 180)
            Line4(q).Y2 = Line4(q).Y2 + XiShu * Sin((270 - JiaoDu4(q)) * Pi / 180)
        Case Is > 270
            Line4(q).X1 = Line4(q).X1 - XiShu * Sin((360 - JiaoDu4(q)) * Pi / 180) / 3
            Line4(q).Y1 = Line4(q).Y1 - XiShu * Cos((360 - JiaoDu4(q)) * Pi / 180) / 3
            Line4(q).X2 = Line4(q).X2 - XiShu * Sin((360 - JiaoDu4(q)) * Pi / 180)
            Line4(q).Y2 = Line4(q).Y2 - XiShu * Cos((360 - JiaoDu4(q)) * Pi / 180)
    End Select
Next q
If Tms(4) = 4 Then
    Tms(4) = 0
    Timer4.Interval = 0
    For i = 0 To 4
        Line4(i).Visible = False
    Next i
End If
End Sub
Private Sub Main_Timer()
For i = 0 To 3
    If EA(i) < 0 Then
        EA(i) = EA(i) + 360
    End If
    With Enemy(i)
        Select Case EA(i)
            Case Is <= 90
                .Left = .Left + ES * Sin(EA(i) * Pi / 180)
                .Top = .Top + ES * Cos(EA(i) * Pi / 180)
            Case 91 To 180
                .Left = .Left + ES * Sin((180 - EA(i)) * Pi / 180)
                .Top = .Top - ES * Cos((180 - EA(i)) * Pi / 180)
            Case 181 To 270
                .Left = .Left - ES * Cos((270 - EA(i)) * Pi / 180)
                .Top = .Top + ES * Sin((270 - EA(i)) * Pi / 180)
            Case Is > 270
                .Left = .Left - ES * Sin((360 - EA(i)) * Pi / 180)
                .Top = .Top - ES * Cos((360 - EA(i)) * Pi / 180)
        End Select
        Select Case True
            Case .Top <= 0
                If EA(i) > 90 And EA(i) < 180 Then
                    EA(i) = 180 - EA(i)
                    GoTo Skip
                End If
                If EA(i) > 270 Then
                    EA(i) = 180 - EA(i)
                End If
            Case .Top >= Form1.Height - .Height
                If EA(i) > 180 And EA(i) < 270 Then
                    EA(i) = 180 - EA(i)
                    GoTo Skip
                End If
                If EA(i) < 90 Then
                    EA(i) = 180 - EA(i)
                End If
            Case .Left <= 0
                If EA(i) > 180 Then
                    EA(i) = 180 + EA(i)
                End If
            Case .Left >= Form1.Width - .Width
                If EA(i) < 180 Then
                    EA(i) = 180 + EA(i)
                    GoTo Skip
                End If
Skip:
        End Select
    End With
Next i
End Sub

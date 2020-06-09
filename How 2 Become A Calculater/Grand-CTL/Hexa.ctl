VERSION 5.00
Begin VB.UserControl Hexa 
   ClientHeight    =   4992
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4440
   ScaleHeight     =   4992
   ScaleWidth      =   4440
   Begin VB.Timer Liner 
      Left            =   3720
      Top             =   1800
   End
   Begin VB.Timer Looker 
      Left            =   1800
      Top             =   2520
   End
   Begin VB.Timer Xiao 
      Enabled         =   0   'False
      Left            =   3600
      Top             =   3840
   End
   Begin VB.Timer Acc 
      Left            =   2520
      Top             =   2880
   End
   Begin VB.Timer Fat 
      Left            =   4560
      Top             =   3480
   End
   Begin VB.Timer Heart 
      Left            =   3000
      Top             =   3120
   End
   Begin VB.Timer Rander 
      Interval        =   1
      Left            =   3360
      Top             =   2400
   End
   Begin VB.Timer Main 
      Left            =   2400
      Top             =   2040
   End
End
Attribute VB_Name = "Hexa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Points
    X As Single
    Y As Single
    Rad As Single
End Type

Const Pi As Single = 3.14159265358979

Const DotSize As Long = 100
Dim Scal As Long
Dim SpinSpeed As Double
Const Interval As Integer = 40
Const FloatSpeed As Double = 1.9

Const DeOY As Long = 2500
Dim OY As Long

Dim Colorful As Boolean
Dim FirstDraw As Boolean
Dim Thick As Single
Dim Shake As Long
Dim OX As Long
Dim Point(1 To 15) As Points
Dim Map(1 To 15)
Dim Rotate As Double
Dim Float As Double
Const DeLook As Single = 0.19
Dim Look As Single
Dim HeartSize As Long
Dim HeartTms As Long
Const HeartSpeed As Long = 1900
Dim Fade As Integer
Dim Fad As Integer
Dim Mad As Boolean
Dim FatSize As Double
Dim FatTms As Long
Const FatItv As Single = 6
Dim Roll As Boolean
Dim RollSpeed As Single
Dim Exhaust As Boolean
Dim AccTms As Double
Dim XiaoTms As Single
Dim LookNotChange As Boolean
Dim DoLine As Boolean
Dim LinerTms As Integer
Dim LookerTms As Integer
Dim SetBack As Boolean
Dim BaColor As Long

Private Sub Acc_Timer()
If AccTms <= 0.29 Then
    AccTms = AccTms + 0.0038
    SpinSpeed = SpinSpeed - AccTms
    RollSpeed = RollSpeed + AccTms / 38
Else
    Acc = False
End If
End Sub

Private Sub Fat_Timer()
Select Case Fat.Tag
    Case "thin"
        FatTms = FatTms + 1
        FatSize = FatSize + 1
        If FatTms >= FatItv Then   '调试
            FatTms = 0
            Fat.Tag = "fat"
        End If
    Case "fat"
        FatTms = FatTms + 1
        FatSize = FatSize - 1
        If FatTms >= FatItv Then   '调试
            FatTms = 0
            Fat.Tag = "willshort"
            FatSize = 0
        End If
    Case "short"
        FatTms = FatTms + 1
        FatSize = FatSize - 1
        If FatTms >= FatItv Then   '调试
            FatTms = 0
            Fat.Tag = "tall"
        End If
    Case "tall"
        FatTms = FatTms + 1
        FatSize = FatSize + 1
        If FatTms >= FatItv Then   '调试
            FatTms = 0
            Fat.Tag = "willthin"
            FatSize = 0
        End If
    Case ""
        Fat.Tag = "willthin"
    Case Else
        FatTms = FatTms + 1
        If FatTms * Fat.Interval / 1000 >= 0.3 Then '调试
            FatTms = 0
            Fat.Tag = Mid(Fat.Tag, 5, 5)
        End If
End Select
End Sub

Private Sub Heart_Timer()
Select Case Heart.Tag
    Case ""
        HeartTms = HeartTms + 1
        If HeartTms * Heart.Interval / 1000 >= 0.3 Or Exhaust Then '调试
            HeartTms = 0
            Heart.Tag = "b"
        End If
    Case "b"
        HeartTms = HeartTms + 1
        HeartSize = HeartSize + HeartSpeed * Heart.Interval / 1000
        If HeartTms / Heart.Interval >= 0.1 Then   '调试
            HeartTms = 0
            Heart.Tag = "d"
        End If
    Case "d"
        HeartTms = HeartTms + 1
        HeartSize = HeartSize - HeartSpeed * Heart.Interval / 1000
        If HeartTms / Heart.Interval >= 0.1 Then   '调试
            HeartTms = 0
            Heart.Tag = ""
            HeartSize = 0
        End If
End Select
End Sub

Private Sub Liner_Timer()
LinerTms = LinerTms + 1
Select Case LinerTms
    Case 2 To 12
        DoLine = True
    Case 27 To 28
        DoLine = True
    Case 31 To 32
        DoLine = True
    Case Is >= 35
        DoLine = True
        Liner = False
        LinerTms = 0
    Case Else
        DoLine = False
End Select
End Sub

Private Sub Looker_Timer()
LookerTms = LookerTms + 1
If LookerTms > 58 Then
    Looker = False
Else
    Look = Pi / 2 + (Cos(LookerTms / 58 * Pi) - 1) * (Pi / 2 - DeLook) / 2
End If
End Sub

Private Sub Main_Timer()
If SetBack Then
    UserControl.BackColor = BaColor
    SetBack = False
End If
Rotate = Rotate + SpinSpeed * Main.Interval / 1000
If Roll Then
    Look = Look - RollSpeed
    If Look < 0 Then
        Look = Look + 2 * Pi
    End If
Else
    If Not LookNotChange Then
        Float = Float + FloatSpeed * Main.Interval / 1000
        If Float > 2 * Pi Then
            Float = Float - 2 * Pi
        End If
        Look = DeLook + 0.1 * Sin(Float)
        OY = DeOY + 0.19 * Scal * Sin(Float)
    End If
End If
DrawWidth = 2
ForeColor = vbWhite
Dim k As Integer
If Fade = 1 Then
    Cls
Else
    If Fade > 1 Then
        For k = Fad + 1 To ScaleHeight Step Fade
            Line (0, k)-(ScaleWidth, k), BackColor
        Next k
        If Mad Then
            Fad = Rnd * Fade
        Else
            Fad = Fad + 4
            Fad = Fad Mod Fade
        End If
    End If
End If
For k = 1 To 14
    Dot DX(Point(k)), DY(Point(k))
Next k
Dot DX(Point(15)), DY(Point(15)), True
Dim kk As Integer
If DoLine Then
    For k = 2 To 15
        If k = 15 Then
            DrawWidth = 1
            For kk = 1 To k - 1
                If Map(k)(kk) = 1 Then
                    If FirstDraw Then
                        ForeColor = vbWhite
                        Line (OX + Scal * DX(Point(kk)), OY + Scal * DY(Point(kk)))-(OX + Scal * DX(Point(k)) + (Rnd - 0.5) * Shake, OY + Scal * DY(Point(k)) + (Rnd - 0.5) * Shake)
                    End If
                    Do While Rnd < Thick
                        If Colorful Then
                            ForeColor = Rnd * vbWhite
                        End If
                        Line (OX + Scal * DX(Point(kk)), OY + Scal * DY(Point(kk)))-(OX + Scal * DX(Point(k)) + (Rnd - 0.5) * Shake, OY + Scal * DY(Point(k)) + (Rnd - 0.5) * Shake)
                    Loop
                End If
            Next kk
        Else
            For kk = 1 To k - 1
                If Map(k)(kk) = 1 Then
                    Line (OX + Scal * DX(Point(kk)), OY + Scal * DY(Point(kk)))-(OX + Scal * DX(Point(k)), OY + Scal * DY(Point(k)))
                End If
            Next kk
        End If
    Next k
End If
End Sub

Private Sub Rander_Timer()
Dim Nm As Integer
Nm = Int(Rnd * 14) + 1
Map(15)(Nm) = 1 - Map(15)(Nm)
End Sub

Private Sub UserControl_Initialize()
With Point(13)
    .X = 0
    .Y = 2
    .Rad = 0
End With
With Point(14)
    .X = 0
    .Y = -2
    .Rad = 0
End With
Dim k As Integer
For k = 1 To 12
    With Point(k)
        .X = Sqr(3)
        .Y = (-1) ^ k
        .Rad = (k \ 2) * Pi / 3
    End With
Next k
OX = ScaleWidth / 2
DrawWidth = 2
SpinSpeed = -0.9
Main.Interval = Interval
Rander.Interval = Interval
Heart.Interval = Interval
Fat.Interval = Interval
Acc.Interval = Interval
Xiao.Interval = Interval
Liner.Interval = Interval
Looker.Interval = Interval
Scal = 1000
DoLine = True
Looker = False
Liner = False
'开头有0
Map(15) = Array(2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
Map(14) = Array(2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 2)
Map(13) = Array(2, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1)
Map(12) = Array(2, 1, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2)
Map(11) = Array(2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 1)
Map(10) = Array(2, 2, 2, 2, 2, 2, 2, 2, 1, 2)
Map(9) = Array(2, 2, 2, 2, 2, 2, 2, 1, 1)
Map(8) = Array(2, 2, 2, 2, 2, 2, 1, 2)
Map(7) = Array(2, 2, 2, 2, 2, 1, 1)
Map(6) = Array(2, 2, 2, 2, 1, 2)
Map(5) = Array(2, 2, 2, 1, 1)
Map(4) = Array(2, 2, 1, 2)
Map(3) = Array(2, 1, 1)
Map(2) = Array(2, 2)
Acc = False
Heart = False
Fat = False
Fade = 1
OY = DeOY
SetMode 1
End Sub

Private Sub Dot(X As Single, Y As Single, Optional HeartAct As Boolean = False)
Dim DS As Long
If HeartAct Then
    DS = DotSize + HeartSize
Else
    DS = DotSize
End If
Line (OX + X * Scal - DS, OY + Y * Scal - DS)-(OX + X * Scal - DS, OY + Y * Scal + DS)
Line (OX + X * Scal - DS, OY + Y * Scal - DS)-(OX + X * Scal + DS, OY + Y * Scal - DS)
Line (OX + X * Scal + DS, OY + Y * Scal + DS)-(OX + X * Scal + DS, OY + Y * Scal - DS)
Line (OX + X * Scal + DS, OY + Y * Scal + DS)-(OX + X * Scal - DS, OY + Y * Scal + DS)
End Sub

Private Function DX(Dian As Points) As Single
DX = Dian.X * Cos(Dian.Rad + Rotate)
If FatSize > 0 Then
    DX = DX * (1 - FatSize / FatItv)
End If
End Function

Private Function DY(Dian As Points) As Single
Dim AX As Single, AY As Single
AY = Dian.Y
AX = -Dian.X * Sin(Dian.Rad + Rotate)
DY = AX * Sin(Look) + AY * Cos(Look)
If FatSize < 0 Then
    If Dian.X <> 0 Then
        DY = DY * (1 + FatSize / FatItv)
    End If
End If
End Function

Public Sub BianXiao()
Xiao = True
HeartSize = 0
End Sub

Private Sub Xiao_Timer()
If XiaoTms <= 7 Then
    XiaoTms = XiaoTms + 0.19
End If
Scal = Scal - XiaoTms
If Scal <= 19 Then
    Scal = 19
    Xiao = False
End If
End Sub

Public Sub SetMode(Mode As Integer)
Select Case Mode
    Case 1
        LookNotChange = True
        Look = Pi / 2
        Rander = False
        DoLine = False
    Case 2
        Rander = False
        Liner = True
    Case 3
        DoLine = True
        Liner = False
        Looker = True
    Case 4
        Look = DeLook
        Rander = False
        DoLine = True
        LookNotChange = False
    Case 5
        DoLine = True
        Look = DeLook
        Rander = True
        FirstDraw = True
    Case 6
        Thick = 0.5
        Shake = 50
        FirstDraw = False
    Case 7
        Thick = 0.9
        Shake = 200
    Case 8
        Colorful = True
    Case 9
        Heart = True
    Case 10
        Fade = 42
    Case 11
        Fat = True
    Case 12
        Fade = 83
        Mad = True
    Case 13
        Fat = False
        Roll = True
        Exhaust = True
        RollSpeed = 0.029
        FatSize = 0
    Case 14
        Mad = False
        Acc = True
    Case 15
        BianXiao
        Heart = False
End Select
End Sub

Public Sub SetBackColor(Color As Long)
BaColor = Color
SetBack = True
End Sub

Public Sub Peace()
Main = False
Cls
End Sub

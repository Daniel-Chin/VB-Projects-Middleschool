VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6156
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6156
   ScaleWidth      =   11340
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Spped 
      Interval        =   1
      Left            =   6840
      Top             =   4080
   End
   Begin VB.Timer Render 
      Interval        =   80
      Left            =   5160
      Top             =   2880
   End
   Begin VB.Timer Main 
      Interval        =   1
      Left            =   1440
      Top             =   1080
   End
   Begin VB.Line rLikeHandsome 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   9
      Index           =   0
      X1              =   7680
      X2              =   7680
      Y1              =   0
      Y2              =   2640
   End
   Begin VB.Line rHandsome 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   9
      Index           =   0
      X1              =   3960
      X2              =   3960
      Y1              =   0
      Y2              =   2640
   End
   Begin VB.Line rType2 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   16
      X1              =   2400
      X2              =   2400
      Y1              =   3960
      Y2              =   6000
   End
   Begin VB.Line rType1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   16
      X1              =   2400
      X2              =   2400
      Y1              =   2640
      Y2              =   3960
   End
   Begin VB.Line rType0 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   16
      X1              =   2400
      X2              =   2400
      Y1              =   0
      Y2              =   2640
   End
   Begin VB.Line rFemalePop 
      BorderColor     =   &H008080FF&
      BorderWidth     =   16
      X1              =   1200
      X2              =   1200
      Y1              =   3600
      Y2              =   5000
   End
   Begin VB.Line rMalePop 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   16
      X1              =   1200
      X2              =   1200
      Y1              =   0
      Y2              =   3600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tpMan
    Type As Byte    '1=Sexy 2=Cool
    Gender As Boolean
    Handsome As Single
    LikeHandsome As Single
    Power As Long
End Type

Private Man(0 To 9999) As tpMan
Private Population As Integer
Const StartPop As Integer = 500
Const MaxPop As Integer = 1000
Const Violate As Single = 0.05
Const Wierd As Single = 0.6
Const T2RATE As Single = 0.5
Private MaleList(0 To 9999) As Integer
Private FemaleList(0 To 9999) As Integer
Private MalePop As Integer, FemalePop As Integer

Private Sub Form_Load()
Dim k As Integer
Randomize
For k = 1 To StartPop
    With Man(k)
        .Gender = Rnd < 0.5
        .Handsome = Rnd
        .LikeHandsome = Rnd
        .Power = Int(Rnd * 3)
        If .Gender Then
            MalePop = MalePop + 1
            MaleList(MalePop) = k
        Else
            FemalePop = FemalePop + 1
            FemaleList(FemalePop) = k
        End If
        If Rnd < Wierd Then
            If Rnd < T2RATE Then
                .Type = 2
            Else
                .Type = 1
            End If
        End If
    End With
Next k
Population = StartPop
For k = 1 To 9
    Load rHandsome(k)
    With rHandsome(k)
        .X1 = rHandsome(k - 1).X1 + 200
        .X2 = .X1
        .Visible = True
    End With
    Load rLikeHandsome(k)
    With rLikeHandsome(k)
        .X1 = rLikeHandsome(k - 1).X1 + 200
        .X2 = .X1
        .Visible = True
    End With
Next k
End Sub

Private Sub Main_Timer()
Fight
Fight
Ahhhh
Ahhhh
Ahhhh
Select Case Sqr(MalePop) * Sqr(FemalePop) * 2
    Case Is <= MaxPop / 3
        Fuck
        Fuck
        Fuck
        Fuck
        Fuck
        Fuck
    Case Is >= MaxPop * 0.66
    Case Else
        Fuck
        Fuck
        Fuck
        Fuck
        Fuck
End Select
End Sub

Private Sub Fight()
Dim Jack As Integer, John As Integer
Debug.Assert MalePop >= 2
Jack = Int(Rnd * MalePop) + 1
John = Int(Rnd * (MalePop - 1)) + 1
If John >= Jack Then
    John = John + 1
End If
If (1 - Man(MaleList(Jack)).Handsome) * Man(MaleList(Jack)).Power < (1 - Man(MaleList(John)).Handsome) * Man(MaleList(John)).Power Then
    Die True, John
Else
    Die True, Jack
End If
End Sub

Private Sub Ahhhh()
If MalePop > FemalePop Then
    Die True, Int(Rnd * MalePop) + 1
Else
    Die False, Int(Rnd * FemalePop) + 1
End If
End Sub

Private Sub Fuck()
Dim Bob(1 To 20) As Integer, Alice As Integer
Debug.Assert MalePop >= 1 And FemalePop >= 1
Dim k As Integer
For k = 1 To 20
    Bob(k) = Int(Rnd * MalePop) + 1
    Bob(k) = MaleList(Bob(k))
Next k
Alice = Int(Rnd * FemalePop) + 1
Alice = FemaleList(Alice)
Dim Love As Single, MaxLove As Single, MyOne As Integer
Select Case Man(Alice).Type
    Case 0
        For k = 1 To Int(15 * Man(Alice).LikeHandsome) + 5
            Love = Man(Alice).LikeHandsome _
                * Man(Bob(k)).Handsome _
                + (1 - Man(Alice).LikeHandsome) _
                * (1 - Man(Bob(k)).Handsome) _
                * Man(Bob(k)).Power
            If Love > MaxLove Then
                MyOne = Bob(k)
                MaxLove = Love
            End If
        Next k
    Case 1
        For k = 1 To 20
            Love = Man(Bob(k)).Handsome
            If Love > MaxLove Then
                MyOne = Bob(k)
                MaxLove = Love
            End If
        Next k
    Case 2
        For k = 1 To 5
            Love = (1 - Man(Bob(k)).Handsome) * Man(Bob(k)).Power
            If Love > MaxLove Then
                MyOne = Bob(k)
                MaxLove = Love
            End If
        Next k
End Select
Population = Population + 1
With Man(Population)
    .Gender = Rnd < 0.5
    .Handsome = (Man(Alice).Handsome + Man(MyOne).Handsome) / 2
    .Handsome = .Handsome + (Violate * 2 * Rnd - Violate)
    If .Handsome < 0 Then .Handsome = 0
    If .Handsome > 1 Then .Handsome = 1
    .LikeHandsome = (Man(Alice).LikeHandsome + Man(MyOne).LikeHandsome) / 2
    .LikeHandsome = .LikeHandsome + (Violate * 2 * Rnd - Violate) / 4
    If .LikeHandsome < 0 Then .LikeHandsome = 0
    If .LikeHandsome > 1 Then .LikeHandsome = 1
    If .Gender Then
    .Power = (Man(Alice).Power + Man(MyOne).Power) / 2
    .Power = .Power + (Violate * 20 * Rnd - Violate)
    If .Power < 0 Then .Power = 0
        MalePop = MalePop + 1
        MaleList(MalePop) = Population
    Else
        FemalePop = FemalePop + 1
        FemaleList(FemalePop) = Population
    End If
    If Rnd < 0.5 Then
        .Type = Man(MyOne).Type
    Else
        .Type = Man(Alice).Type
    End If
End With
End Sub

Private Sub Die(Gender As Boolean, Index As Integer)
If Gender Then
    MaleList(Index) = MaleList(MalePop)
    MalePop = MalePop - 1
Else
    FemaleList(Index) = FemaleList(FemalePop)
    FemalePop = FemalePop - 1
End If
Man(Index) = Man(Population)
Population = Population - 1
End Sub

Private Sub Analyze()
Dim TypeNum(0 To 2) As Integer
Dim Hand(0 To 9) As Integer, LikeHand(0 To 9) As Integer
rMalePop.Y2 = Me.ScaleHeight * MalePop / MaxPop
rFemalePop.Y1 = rMalePop.Y2
rFemalePop.Y2 = rMalePop.Y2 + Me.ScaleHeight * FemalePop / MaxPop
Dim k As Integer, B As Integer
For k = 1 To Population
    With Man(k)
        B = Int(.Handsome * 10)
        If B = 10 Then B = 9
        Hand(B) = Hand(B) + 1
        B = Int(.LikeHandsome * 10)
        If B = 10 Then B = 9
        LikeHand(B) = LikeHand(B) + 1
        TypeNum(.Type) = TypeNum(.Type) + 1
    End With
Next k
For k = 0 To 9
    With rHandsome(k)
        .Y2 = Me.ScaleHeight * Hand(k) / Population * 2
    End With
    With rLikeHandsome(k)
        .Y2 = Me.ScaleHeight * LikeHand(k) / Population * 2
    End With
Next k
If TypeNum(0) = 0 Then Stop
If TypeNum(1) = 0 Then Stop
If TypeNum(2) = 0 Then Stop
rType0.Y2 = Me.ScaleHeight * TypeNum(0) / MaxPop
rType1.Y1 = rType0.Y2
rType1.Y2 = rType0.Y2 + Me.ScaleHeight * TypeNum(1) / MaxPop
rType2.Y1 = rType1.Y2
rType2.Y2 = rType1.Y2 + Me.ScaleHeight * TypeNum(2) / MaxPop
Me.Caption = Man(1).Power
End Sub

Sub Disp(Index As Integer)
Debug.Print IIf(Man(Index).Gender, "ÄÐ", "Å®")
Debug.Print Man(Index).Handsome
Debug.Print Man(Index).LikeHandsome
Debug.Print Man(Index).Type
Debug.Print Man(Index).Power
Debug.Print "disp(" & Index + 1 & ")"
End Sub

Sub what()
Dim k
For k = 1 To Population
    Debug.Assert Man(k).Type <> 0
Next k
End Sub

Private Sub Render_Timer()
Analyze
End Sub

Private Sub Spped_Timer()
Dim k As Integer
For k = 1 To 8
    Main_Timer
Next k
End Sub

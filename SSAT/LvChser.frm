VERSION 5.00
Begin VB.Form LvChser 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "我去冒险"
   ClientHeight    =   5916
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11484
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   22.2
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LvChser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5916
   ScaleWidth      =   11484
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Main 
      Left            =   6960
      Top             =   3600
   End
   Begin VB.Line G 
      BorderWidth     =   16
      Index           =   0
      X1              =   -96
      X2              =   4000
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bt"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   25.8
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   516
      Index           =   0
      Left            =   5472
      MouseIcon       =   "LvChser.frx":0CCA
      TabIndex        =   0
      Top             =   2664
      Width           =   552
   End
   Begin VB.Line S 
      Index           =   0
      Visible         =   0   'False
      X1              =   4440
      X2              =   6000
      Y1              =   3120
      Y2              =   2520
   End
End
Attribute VB_Name = "LvChser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Pi As Double = 3.14159265358979
Dim nM As Boolean
Dim k As Long, KK As Integer
Dim SX As Long, Sy As Long
Const FxLays As Integer = 6
Dim FxNum(0 To 6) As Integer
Dim GB As Integer
Dim TopLv As Integer
Dim Tms As Integer
Const WindV As Integer = 466

Private Sub Bt_Click(Index As Integer)
If Index >= 0 And Index <= 23 Then
    If GB = Index Then
        Select Case Bt(GB).Caption
            Case "Lv1"
                Lv1.Show
            Case "Lv2"
                Lv2.Show
            Case "Lv3"
                Lv3.Show
            Case "Lv3.5"
                Lv3_5.Show
            Case "Lv4"
                Lv4.Show
            Case "Lv5"
                Lv5.Show
            Case "Lv6"
                Lv6.Show
            Case "Lv7"
                Lv7.Show
            Case "Lv8"
                Lv8.Show
            Case "Lv8.5"
                Lv8_5.Show
            Case "Lv9"
                Lv9.Show
            Case "Lv10"
                Lv10.Show
            Case "Lv11"
                Lv11.Show
            Case "Lv12"
                Lv12.Show
            Case "Lv13"
                Lv13.Show
            Case "Lv14"
                Lv14.Show
            Case "Lv14.5"
                Lv14_5.Show
            Case "Lv15"
                Lv15.Show
            Case "LvBoss"
                LvBoss.Show
            Case "偷西瓜"
                LvDog.Show
            Case "迷宫"
                LvMaze.Show
            Case "高等数学"
                LvMath.Show
            Case "第一人称FPS"
                LvFPS.Show
            Case "异"
                LvBi.Show
        End Select
        nM = True
        Unload Me
    Else
        If TopLv >= Index Then
            If Bt(GB).Tag = "b" Then
                Bt(GB).ForeColor = vbBlue
                Bt(GB).BackStyle = 0
            Else
                Bt(GB).ForeColor = 0
                Bt(GB).BackStyle = 0
            End If
            GB = Index
            Bt(GB).ForeColor = vbWhite
            Bt(GB).BackStyle = 1
            Bt(GB).BackColor = 0
        End If
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 38 'Up
        Bt_Click GB + 1
    Case 40 'Down
        Bt_Click GB - 1
    Case 13
        G(0).BorderColor = vbBlue
        G(1).BorderColor = vbBlue
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 27
        Unload Me
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    G(0).BorderColor = 0
    G(1).BorderColor = 0
    Bt_Click GB
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Randomize
Main_Timer
FxNum(0) = 1
Dim N As Integer
N = 36
For k = 1 To FxLays
    FxNum(k) = FxNum(k - 1) + N
    N = Int(N * 0.36)
Next k
Main.Interval = 12
Dim KK As Integer
For KK = 1 To FxLays
    For k = FxNum(KK - 1) To FxNum(KK) - 1
        Load S(k)
        With S(k)
            .BorderColor = RGB(0, 0, Int(Rnd() ^ 2 * 256))
            .BorderWidth = KK + 1
            .X1 = Rnd * Width
            .X2 = .X1
            .Y1 = Rnd * Height
            .Y2 = .Y1
            .Visible = True
        End With
    Next k
Next KK
Dim BtC
BtC = Array("Lv1", "Lv2", "Lv3", "Lv3.5", "Lv4", "Lv5", "Lv6", "Lv7", "Lv8", "Lv8.5", "Lv9", "Lv10", "Lv11", "Lv12", "Lv13", "Lv14", "Lv14.5", "Lv15", "LvBoss", "偷西瓜", "迷宫", "高等数学", "第一人称FPS", "异")
For k = 0 To 23
    With Bt(k)
    If k <> 0 Then
        Load Bt(k)
        .Top = Bt(k - 1).Top - .Height
        .Visible = True
    End If
    .Caption = BtC(k)
    End With
Next k
Load G(1)
With G(1)
    .X2 = Me.Width + 100
    .X1 = Me.Width - G(0).X2
    .Visible = True
    .ZOrder 0
End With
'调整默认GB、TopLv
TopLv = Int(GetSetting("iGlope", "woqu", "lv", "1")) - 1
GB = TopLv
If TopLv = 18 Then
    If GetSetting("iGlope", "woqu", "sq1", 0) = 1 Then
        TopLv = 23
        Bt(18).Tag = "b"
        Bt(18).ForeColor = vbBlue
        For k = 19 To 23
            If GetSetting("iGlope", "woqu", "j" & k - 18, 0) = 1 Then
                Bt(k).ForeColor = vbBlue
                Bt(k).Tag = "b"
            Else
                Bt(k).ForeColor = 0
            End If
        Next k
    Else
        For k = 19 To 23
            Bt(k).ForeColor = &HCCCCCC
        Next k
    End If
    For k = 0 To 17
        Bt(k).ForeColor = vbBlue
        Bt(k).Tag = "b"
    Next k
Else
    For k = TopLv + 1 To 23
        Bt(k).ForeColor = &HCCCCCC
    Next k
    If TopLv > 0 Then
        For k = 0 To TopLv
            Bt(k).ForeColor = vbBlue
            Bt(k).Tag = "b"
        Next k
    End If
End If
Bt(GB).ForeColor = vbWhite
Bt(GB).BackStyle = 1
Bt(GB).BackColor = 0
End Sub

Private Sub Main_Timer()
Sy = Sy * 0.96
Tms = (Tms + 1) Mod WindV * 2
If Tms Mod 6 = 0 Then
    Dim Ang As Double
    Ang = Tms / WindV * Pi - Pi
    SX = Sin(Ang) * 66
End If
If Tms Mod 2 = 0 Then
    Dim GBP As Integer
    GBP = G(0).Y1 - Bt(GB).Top - Bt(GB).Height * 0.4
    If Abs(GBP) > 6 Then
        GBP = GBP / 6
    End If
    If Abs(GBP) > Abs(Sy) Then
        Sy = GBP
    End If
    For k = 0 To 23
        Bt(k).Top = Bt(k).Top + GBP
    Next k
End If
KK = KK Mod FxLays
KK = KK + 1
For k = FxNum(KK - 1) To FxNum(KK) - 1
    With S(k)
        .X1 = .X1 + SX * KK * 0.36
        If .X1 > Width + 100 Then .X1 = -100: .Y1 = Rnd * Height
        If .X1 < -100 Then .X1 = Width + 100: .Y1 = Rnd * Height
        .Y1 = .Y1 + Sy * KK * 0.36
        If .Y1 > Height + 100 Then .Y1 = -100: .X1 = Rnd * Width
        If .Y1 < -100 Then .Y1 = Height + 100: .X1 = Rnd * Width
        .X2 = .X1 - SX * KK
        .Y2 = .Y1 - Sy * KK
    End With
Next k
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not nM Then
    Shell App.Path & "\" & App.EXEName, vbNormalFocus
    End
End If
End Sub


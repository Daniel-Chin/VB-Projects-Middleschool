VERSION 5.00
Begin VB.Form M 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Matrix"
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9036
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   9036
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Main 
      Left            =   6120
      Top             =   3600
   End
   Begin VB.Timer Loader 
      Interval        =   1
      Left            =   3600
      Top             =   2520
   End
   Begin VB.Shape Fuzhu 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   5760
      Index           =   1
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   5760
   End
   Begin VB.Shape Fuzhu 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   2880
      Index           =   0
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   5760
      Width           =   2880
   End
   Begin VB.Shape P 
      FillStyle       =   0  'Solid
      Height          =   372
      Index           =   0
      Left            =   3720
      Top             =   1800
      Visible         =   0   'False
      Width           =   252
   End
End
Attribute VB_Name = "M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Const xS As Long = 96
Dim M(0 To 89, 0 To 89) As Integer
Dim Loaded As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
If Loaded Then
    Select Case KeyAscii
        Case 27     'ESC
            End
        Case 32     'Space
            RenewAll
            Main.Interval = 1 - Main.Interval
        Case 18     'Ctrl+r
            RS
    End Select
End If
End Sub

Private Sub Form_Load()
Randomize
With Me
.Top = 0
.Left = 0
.Width = Screen.Width
.Height = Screen.Width
End With
End Sub

Private Sub Loader_Timer()
For j = 0 To 511
i = i + 1
If i = 8101 Then
    Loader.Interval = 0
    MsgBox "Matrix loaded.", vbInformation + vbOKOnly, "Information"
    Loaded = True
    Exit Sub
End If
Load P(i)
With P(i)
    i = i - 1
    .Top = (i Mod 90) * xS
    .Left = (i \ 90) * xS
    .Height = xS
    .Width = .Height
    .BorderWidth = 1
    .Visible = True
End With
i = i + 1
Next j
End Sub

Private Function Zhuan(X As Integer, Y As Integer) As Integer
    Zhuan = X * 90 + Y + 1
End Function

Private Sub Main_Timer()
Dim MyX As Integer, MyY As Integer, TarX As Integer, TarY As Integer
For jj = 0 To 25    '时间进行速度
For j = 0 To 2047
    MyX = Int(Rnd() * 90)
    MyY = Int(Rnd() * 90)
    If M(MyX, MyY) <> 0 Then
        Exit For
    End If
    If j = 2047 Then
        MsgBox "灭绝了。", vbOKOnly + vbCritical, "Matrix"
        Main.Interval = 0
        Exit Sub
    End If
Next j
TarX = MyX + Int(Rnd() * 3) - 1
TarY = MyY + Int(Rnd() * 3) - 1
If TarX > 89 Or TarX < 0 Then TarX = MyX
If TarY > 89 Or TarY < 0 Then TarY = MyY
If Sqr((TarX - 45) ^ 2 + (TarY - 45) ^ 2) < 30 Or Sqr((TarX - 75) ^ 2 + (TarY - 75) ^ 2) < 15 Then '检测水井
    '有水
Else
    If Rnd() < 0.5 Then
        M(MyX, MyY) = 0 '衰弱系数
        GoTo JieShu     '直接结束
    End If
End If
If TarX = MyX And TarY = MyY Then GoTo JieShu
If M(TarX, TarY) = 0 Then   '目标无生物
    M(TarX, TarY) = M(MyX, MyY)
    M(MyX, MyY) = 0
Else
    If M(TarX, TarY) Mod 2 = M(MyX, MyY) Mod 2 Then '相同性别
        If M(MyX, MyY) > M(TarX, TarY) Then '战胜
            M(TarX, TarY) = M(MyX, MyY)
            M(MyX, MyY) = 0
        Else    '战败
            If Rnd() < 0.01 Then     '逃脱概率
                M(MyX, MyY) = 0
            End If
        End If
    Else    '不同性别
        Dim Son As Integer
        Son = M(MyX, MyY) \ 2 + M(TarX, TarY) \ 2
        Son = Son + Int(Rnd() * 101) - 80   '变异系数
        If Son < 0 Then Son = 0 '生育失败
        For j = -10 To 10
            If MyX + j > 89 Or MyX + j < 0 Then Exit For
            If M(MyX + j, MyY) = 0 Then
                M(MyX + j, MyY) = Son   '生子
                Renew MyX + j, MyY
                Exit For
            End If
        Next j
    End If
End If
JieShu: Renew MyX, MyY
        Renew TarX, TarY
Next jj
End Sub

Private Sub Renew(X As Integer, Y As Integer)
If M(X, Y) Mod 2 = 0 Then   '男性
    P(Zhuan(X, Y)).FillColor = RGB(0, 0, M(X, Y))
Else    '女性
    P(Zhuan(X, Y)).FillColor = RGB(M(X, Y), 0, 0)
End If
End Sub

Private Sub RS()
For j = 10 To 80
    For q = 10 To 80
        M(j, q) = IIf(Rnd() < 0.3, Int(Rnd() * 150), 0)
    Next q
Next j
End Sub

Private Sub Bomb()
For j = 35 To 55
    For q = 40 To 50
        M(j, q) = 0
    Next q
Next j
RenewAll
End Sub

Private Sub RenewAll()
Dim j As Integer, q As Integer
    For j = 0 To 89
        For q = 0 To 89
            Renew j, q
        Next q
    Next j
End Sub

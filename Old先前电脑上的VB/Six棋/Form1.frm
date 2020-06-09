VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "蜂窝五子棋"
   ClientHeight    =   5568
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8268
   ForeColor       =   &H8000000E&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5568
   ScaleWidth      =   8268
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Reseter 
      Left            =   3600
      Top             =   2640
   End
   Begin VB.Label Lab 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 白方"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   612
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1452
   End
   Begin VB.Label Ge 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   500
      Index           =   1
      Left            =   3840
      MouseIcon       =   "Form1.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   720
      Width           =   500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v(1 To 162) As Byte
Dim PreGe As Byte
Dim P As Boolean
Dim Paused As Boolean
Dim R As Integer
Dim Pre As Integer
Dim YiHui As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27     'ESC
            End
        Case 32     'SPACE
            If MsgBox("重置棋盘？", vbYesNo + vbQuestion, "蜂窝询问您：") - 7 Then
                Reset
            End If
        Case 26     'Ctrl + Z
        If YiHui = False Then
            Paused = False
            YiHui = True
            v(Pre) = 0
            P = Not P
            With Lab
                .Caption = "悔棋"
                .BackColor = vbWhite - .BackColor
                .ForeColor = vbWhite - .ForeColor
            End With
            Ge_MouseMove Pre, 0, 0, 0, 0
            For i = 1 To 162
                Ge(i).ForeColor = IIf(P, 0, vbWhite)
            Next i
        End If
    End Select
End Sub

Private Sub Form_Load()
    Lab.Caption = " 白方"
    PreGe = 1
    Ge(1).BackColor = RGB(150, 150, 150)
    With Me
        .Top = 0
        .Left = 0
    End With
    With Screen
        Me.Height = .Height
        Me.Width = .Width
    End With
    Dim i As Byte
    Dim X As Long, Y As Long
    For i = 2 To 162    '一共162格
        If i Mod 25 = 13 Or i Mod 25 = 1 Then   '决定是否换行
            X = Ge(1).Left  '布局
            Y = Ge(i - 1).Top + Ge(i - 1).Height * 1.2
            If i Mod 25 = 13 Then   '是不是“缩进去”的一行
                X = X - Ge(1).Width / 2 * 1.2   '如果是，格子往右移0.5个单位
            End If
        Else
            X = Ge(i - 1).Left + Ge(i - 1).Width * 1.2
            Y = Ge(i - 1).Top
        End If
        Load Ge(i)  '加载控件
        Ge(i).Left = X
        Ge(i).Top = Y
        Ge(i).Visible = True
    Next i
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case v(PreGe)
        Case 0
            Ge(PreGe) = ""
        Case 1
            Ge(PreGe).BackColor = vbBlack
            Ge(PreGe) = ""
        Case 2
            Ge(PreGe).BackColor = vbWhite
            Ge(PreGe) = ""
    End Select
End Sub

Private Sub Ge_Click(Index As Integer)
    Dim i As Byte
    i = Index
    If Paused Then
        Exit Sub
    End If
    If v(Index) = 0 Then
        YiHui = False
        Pre = Index
        v(Index) = P + 2
        Ge_MouseMove Index, 0, 0, 0, 0
        P = Not P
        If P Then
            Lab.Caption = " 黑方"
            Lab.BackColor = vbBlack
            Lab.ForeColor = vbWhite
            For q = 1 To 162
                Ge(q).ForeColor = vbBlack
            Next q
        Else
            Lab.Caption = " 白方"
            Lab.BackColor = vbWhite
            Lab.ForeColor = vbblavk
            For q = 1 To 162
                Ge(q).ForeColor = vbWhite
            Next q
        End If
        If Win(i) Then
            For q = 1 To 162
                Ge(q).ForeColor = vbRed
            Next q
            If P Then
                Lab = "白胜"
                Lab.BackColor = vbGreen
                Lab.ForeColor = vbWhite
            Else
                Lab = "黑胜"
                Lab.BackColor = vbGreen
                Lab.ForeColor = vbBlack
            End If
            v(Index) = 3
            Paused = True
            Ge_MouseMove Index, 0, 0, 0, 0
        End If
    End If
End Sub

Private Sub Ge_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case v(Index)
        Case 0
            Ge(Index) = "※"
        Case 1
            Ge(Index).BackColor = RGB(30, 30, 30)
            Ge(PreGe) = "※"
        Case 2
            Ge(Index).BackColor = RGB(220, 220, 220)
            Ge(PreGe) = "※"
        Case 3
            Ge(Index).BackColor = RGB(0, 200, 0)
            Ge(PreGe) = "※"
    End Select
    If Index = PreGe Then
        Exit Sub
    End If
    Select Case v(PreGe)
        Case 0
            Ge(PreGe) = ""
            Ge(Index).BackColor = RGB(150, 150, 150)
        Case 1
            Ge(PreGe).BackColor = vbBlack
            Ge(PreGe) = ""
        Case 2
            Ge(PreGe).BackColor = vbWhite
            Ge(PreGe) = ""
        Case 3
            Ge(PreGe).BackColor = vbGreen
            Ge(PreGe) = ""
    End Select
    PreGe = Index
End Sub

Private Function Win(GeZi As Byte) As Boolean
        Dim My As Byte
        Dim pSt As Byte
        Dim S As Byte
        My = v(GeZi)
        S = 1
        pSt = GeZi
hE1:    If pSt Mod 25 = 12 Or pSt Mod 25 = 0 Then
            GoTo sKh1
        End If
        pSt = pSt + 1
        If v(pSt) = My Then
            S = S + 1
            GoTo hE1
        End If
sKh1:   pSt = GeZi
hE2:    If pSt Mod 25 = 1 Or pSt Mod 25 = 13 Then
                GoTo sKh2
        End If
        pSt = pSt - 1
        If v(pSt) = My Then
            S = S + 1
            GoTo hE2
        End If
sKh2:   If S >= 5 Then
            Win = True
            Exit Function
        End If
        S = 1
''''''''
        pSt = GeZi
nA1:    If pSt Mod 25 = 0 Or pSt > 150 Then
            GoTo sKn1
        End If
        pSt = pSt + 13
        If v(pSt) = My Then
            S = S + 1
            GoTo nA1
        End If
sKn1:   pSt = GeZi
nA2:    If pSt Mod 25 = 13 Or pSt < 13 Then
                GoTo sKn2
        End If
        pSt = pSt - 13
        If v(pSt) = My Then
            S = S + 1
            GoTo nA2
        End If
sKn2:   If S >= 5 Then
            Win = True
            Exit Function
        End If
        S = 1
''''''''
        pSt = GeZi
pE1:    If pSt Mod 25 = 0 Or pSt > 150 Then
            GoTo sKp1
        End If
        pSt = pSt + 12
        If v(pSt) = My Then
            S = S + 1
            GoTo pE1
        End If
sKp1:   pSt = GeZi
pE2:    If pSt Mod 25 = 13 Or pSt < 13 Then
                GoTo sKp2
        End If
        pSt = pSt - 12
        If v(pSt) = My Then
            S = S + 1
            GoTo pE2
        End If
sKp2:   If S >= 5 Then
            Win = True
            Exit Function
        End If
End Function

Private Sub Reset()
    Paused = True
    P = (Rnd() < 0.5)
    If P Then
        Lab.Caption = " 黑方"
        Lab.BackColor = 0
        Lab.ForeColor = vbWhite
    Else
        Lab.Caption = " 白方"
        Lab.BackColor = vbWhite
        Lab.ForeColor = 0
    End If
    For i = 1 To 162
        v(i) = 0
    Next i
    Reseter.Interval = 5
    R = 0
End Sub

Private Sub Reseter_Timer()
    R = R + 1
    Ge_MouseMove R, 0, 0, 0, 0
    Ge_MouseMove R + 54, 0, 0, 0, 0
    Ge_MouseMove R + 108, 0, 0, 0, 0
    Ge(R).ForeColor = vbGreen
    If R = 54 Then
        R = 0
        Reseter.Interval = 0
        Paused = False
            For i = 1 To 162
                Ge(i).ForeColor = IIf(P, vbwihte, 0)
            Next i
    End If
End Sub

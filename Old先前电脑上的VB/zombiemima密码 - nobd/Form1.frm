VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Jack - 登录"
   ClientHeight    =   3048
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7548
   LinkTopic       =   "Form1"
   ScaleHeight     =   3048
   ScaleWidth      =   7548
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer UnBlocker 
      Left            =   7080
      Top             =   960
   End
   Begin VB.Timer Checker 
      Left            =   6600
      Top             =   240
   End
   Begin VB.Timer Blocker 
      Left            =   6600
      Top             =   840
   End
   Begin VB.TextBox T 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Index           =   6
      Left            =   6708
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1560
      Width           =   612
   End
   Begin VB.TextBox T 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Index           =   5
      Left            =   5628
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1554
      Width           =   612
   End
   Begin VB.TextBox T 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Index           =   4
      Left            =   4548
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1554
      Width           =   612
   End
   Begin VB.TextBox T 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Index           =   3
      Left            =   3468
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1554
      Width           =   612
   End
   Begin VB.TextBox T 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Index           =   2
      Left            =   2388
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1554
      Width           =   612
   End
   Begin VB.TextBox T 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Index           =   1
      Left            =   1308
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1554
      Width           =   612
   End
   Begin VB.TextBox T 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Index           =   0
      Left            =   228
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   612
   End
   Begin VB.Label lB 
      Alignment       =   2  'Center
      Caption         =   "PassWord"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   1626
      TabIndex        =   6
      Top             =   240
      Width           =   4296
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Blocked As Boolean
Dim BlockerTms As Byte, UnBlockerTms As Byte
Dim CheckerTms As Byte

Private Sub Blocker_Timer()
    BlockerTms = BlockerTms + 1
    Dim a As String
    a = "BLOCKED"
    Select Case BlockerTms
        Case 1 To 7
            T(BlockerTms - 1) = Mid(a, BlockerTms, 1)
        Case 8
            Dim k As Byte
            For k = 0 To 6
                T(k).ForeColor = vbRed
            Next k
        Case 20
            For k = 0 To 6
                T(k) = ""
                T(k).ForeColor = vbBlack
            Next k
            T(0).SetFocus
            BlockerTms = 0
            Blocker.Interval = 0
    End Select
End Sub

Private Sub Checker_Timer()
    CheckerTms = CheckerTms + 1
    Select Case CheckerTms
        Case 1 To 20
            Dim k As Byte
            For k = 0 To 6
                T(k) = Chr(Int(Rnd() * 75) + 48)
            Next k
        Case 33
            Dim a As String
            a = "icorect"
            For k = 0 To 6
                T(k).ForeColor = vbRed
                T(k) = Mid(a, k + 1, 1)
            Next k
        Case 50
            For k = 0 To 6
                T(k).ForeColor = vbBlack
                T(k) = ""
            Next k
            T(0).SetFocus
            CheckerTms = 0
            Checker.Interval = 0
    End Select
End Sub

Private Sub Form_Load()
    Randomize
    Show
    T(0).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "攻略：打开程序，方向键按上下左右，再按四下大写锁定，回车，按住ESC等一分钟。"
End Sub

Private Sub T_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim Legel As Boolean
    Legel = True
    If KeyCode = 27 Then
        If Second(Now) <= 1 Then
            If T(0).ForeColor = vbBlue Then
                Call cRt
            End If
        End If
    End If
    Select Case True
        Case Index = 0 And KeyCode = 38
            T(Index) = "I"
        Case Index = 1 And KeyCode = 40
            T(Index) = "'"
        Case Index = 2 And KeyCode = 37
            T(Index) = "m"
        Case Index = 3 And KeyCode = 39
            T(Index) = "_"
        Case Index = 4 And KeyCode = 20
            T(Index) = "Q"
        Case Index = 5 And KeyCode = 20
            T(Index) = "n"
        Case Index = 6 And KeyCode = 20
            T(Index) = "f"
        Case Else
            Legel = False
    End Select
    If Legel And Index <= 5 Then
        T(Index + 1).SetFocus
    End If
End Sub

Private Sub T_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Legel As Boolean
    If Checker.Interval = 0 And Blocker.Interval = 0 Then
        Select Case KeyAscii
            Case 48 To 57
                Legel = True
            Case 97 To 122
                Legel = True
            Case 65 To 90
                Legel = True
            Case 32
                Legel = True
        End Select
        If Legel Then
            T(Index) = Chr(Int(Rnd() * 75) + 48)
            If Index <= 5 Then
                T(Index + 1).SetFocus
            End If
        End If
        If KeyAscii = 13 Then
            Dim k As Byte, a As String
            For k = 0 To 6
                a = a & T(k)
            Next k
            If a = "I'm_Qnf" Then
                a = "Correct"
                For k = 0 To 6
                    T(k).ForeColor = vbBlue
                    T(k) = Mid(a, k + 1, 1)
                Next k
                T(0).SetFocus
            Else
                If Blocked Then
                    Blocker.Interval = 50
                Else
                    Blocked = True
                    UnBlocker.Interval = 200
                    Checker.Interval = 50
                End If
            End If
        End If
    End If
    If KeyAscii = 8 Then    '退格
        If Index >= 1 Then
            T(Index - 1).SetFocus
            T(Index) = ""
        Else
            T(Index) = ""
        End If
    End If
End Sub

Private Sub UnBlocker_Timer()
    UnBlockerTms = UnBlockerTms + 1
    If UnBlockerTms = 30 Then
        Blocked = False
        UnBlockerTms = 0
        UnBlocker.Interval = 0
    End If
End Sub

Private Sub cRt()
    MsgBox "恭喜你真正通过了Jack安保检查。", vbInformation + vbOKOnly, "欢迎"
    End
End Sub

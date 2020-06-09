VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Rand"
   ClientHeight    =   5832
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   22.2
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5832
   ScaleWidth      =   10740
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Fixer 
      Enabled         =   0   'False
      Interval        =   36
      Left            =   9000
      Top             =   3720
   End
   Begin VB.Timer Ac 
      Enabled         =   0   'False
      Interval        =   36
      Left            =   7200
      Top             =   2520
   End
   Begin VB.Timer Main 
      Interval        =   66
      Left            =   8040
      Top             =   3360
   End
   Begin VB.Frame Frame1 
      Height          =   2412
      Left            =   3120
      TabIndex        =   1
      Top             =   2520
      Width           =   4452
      Begin VB.Label S 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "66"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   66
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1344
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   16
      X1              =   5400
      X2              =   5400
      Y1              =   2040
      Y2              =   5400
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   372
      Index           =   13
      Left            =   3960
      Shape           =   2  'Oval
      Top             =   3960
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Random Number Producer! "
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1152
      TabIndex        =   0
      Top             =   600
      Width           =   8436
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const KD As Integer = 5
Dim V As Long
Dim Tms As Integer

Private Sub Ac_Timer()
V = V + 16
V = V * 1.1
With Frame1
    .Left = .Left - V
    Do While .Left <= -Me.Width / KD
        .Left = .Left + Me.Width / KD
        For k = 0 To KD - 1
            S(k) = S(k + 1)
        Next k
        S(KD) = RN
    Loop
End With
If V >= 1066 Then
    V = 1066
    Ac = False
    Main = True
End If
End Sub

Private Sub Fixer_Timer()
If Tms Mod 2 = 0 Then
    V = V - 2
    Frame1.Left = Frame1.Left + V
    If Frame1.Left <= -Me.Width / KD Then
        Frame1.Left = -Me.Width / KD
    Else
        Exit Sub
    End If
Else
    V = V + 2
    Frame1.Left = Frame1.Left + V
    If Frame1.Left >= 0 Then
        Frame1.Left = 0
    Else
        Exit Sub
    End If
End If
Fixer = False
If Tms Mod 2 = 0 Then
    With S(3)
        .BackColor = 0
        .ForeColor = &HFFEECC
    End With
Else
    With S(2)
        .BackColor = 0
        .ForeColor = &HFFEECC
    End With
End If
Frame1.BackColor = vbWhite
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Ac Or Fixer Or Main Then
Else
    Ac = True
    V = 0
    Tms = Tms + 1
    With S(2)
        .BackColor = vbWhite
        .ForeColor = 0
    End With
    With S(3)
        .BackColor = vbWhite
        .ForeColor = 0
    End With
    Frame1.BackColor = &H661111
End If
End Sub

Private Sub Main_Timer()
With Frame1
    .Left = .Left - V
    Do While .Left <= -Me.Width / KD
        .Left = .Left + Me.Width / KD
        For k = 0 To KD - 1
            S(k) = S(k + 1)
        Next k
        S(KD) = RN
    Loop
End With
V = V * 0.99 - 1
If Tms = 2 And V <= 10 Then
    V = -10
    Main = False
    Fixer = True
End If
If V = 0 Then
    Main = False
    Fixer = True
End If
End Sub

Private Sub Form_Load()
Main.Interval = 36
Main = False
Randomize
Me.BackColor = &H330000
Label1.BackColor = &HFFEECC
Label1.ForeColor = 0
With Frame1
    .BackColor = &H661111
    .Left = 0
    .Width = Me.Width / KD * (KD + 1)
End With
For k = 0 To KD
    If k <> 0 Then Load S(k)
    With S(k)
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Caption = RN
        .Left = Me.Width / KD * k
        .Width = Me.Width / KD
        .AutoSize = True
        .Visible = True
    End With
Next k
End Sub

Function RN() As String
RN = Format(Int(Rnd() * 47) + 1, "00")
End Function

Private Sub S_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1_MouseUp Button, Shift, X, Y
End Sub

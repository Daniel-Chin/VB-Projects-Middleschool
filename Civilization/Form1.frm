VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   10344
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   10344
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   3240
   End
   Begin VB.TextBox Focusa 
      Height          =   372
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Here"
      Top             =   4200
      Width           =   972
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   3840
      Top             =   2760
   End
   Begin VB.Shape LZ 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      Height          =   492
      Left            =   1920
      Top             =   1560
      Width           =   7452
   End
   Begin VB.Shape ZZ 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   3
      Height          =   492
      Left            =   960
      Top             =   3480
      Width           =   7452
   End
   Begin VB.Label Cap 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   4080
      TabIndex        =   5
      Top             =   1560
      Width           =   2412
   End
   Begin VB.Label Blue 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3972
      Index           =   1
      Left            =   8526
      TabIndex        =   3
      Top             =   600
      Width           =   732
   End
   Begin VB.Label Brown 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3972
      Index           =   1
      Left            =   7560
      TabIndex        =   2
      Top             =   600
      Width           =   732
   End
   Begin VB.Label Blue 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3972
      Index           =   0
      Left            =   2046
      TabIndex        =   1
      Top             =   600
      Width           =   732
   End
   Begin VB.Label Brown 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3972
      Index           =   0
      Left            =   1086
      TabIndex        =   0
      Top             =   600
      Width           =   732
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Con(0 To 1) As Boolean
Dim His(0 To 1, 0 To 999) As Integer
Const Delay As Single = 1.9
Const sC As Long = 900
Dim HisTop As Integer
Dim BlueAcc(0 To 1) As Single
Dim BH(0 To 1) As Single

Private Sub Focusa_KeyDown(KeyCode As Integer, Shift As Integer)
'90 88 97 98
Select Case KeyCode
    Case 90
        Con(0) = True
    Case 88
        Con(0) = False
    Case 97
        Con(1) = True
    Case 98
        Con(1) = False
End Select
End Sub

Private Sub Form_Load()
HisTop = Delay * 1000 / Timer.Interval
Dim P As Integer
For P = 0 To 1
    Blue(P).Height = 0
    Brown(P).Height = 0
Next P
Me.ScaleHeight = sC
ZZ.Top = -999
LZ.Top = -999
End Sub

Private Sub Timer_Timer()
Dim P As Integer, k As Integer
For P = 0 To 1
    If Con(P) Then
        His(P, HisTop) = 1
    Else
        His(P, HisTop) = 2
    End If
    If His(P, 0) > 0 Then
        If His(P, 0) = 1 Then
            Brown(P).Height = Brown(P).Height + 3
            If Brown(P).Height + Brown(P).Top > ZZ.Top + ZZ.Height Then
                ZZ.Top = Brown(P).Height - ZZ.Height + Brown(P).Top
            End If
        Else
            BlueAcc(P) = BlueAcc(P) + 0.06
            BH(P) = BH(P) + 0.5
        End If
        BH(P) = BH(P) + BlueAcc(P)
        Blue(P).Height = BH(P)
        If Blue(P).Height + Blue(P).Top > LZ.Top + LZ.Height Then
            LZ.Top = Blue(P).Height - LZ.Height + Blue(P).Top
        End If
        If Blue(P).Height + Blue(P).Top < LZ.Top Then
            Blue(1 - P).BackColor = vbWhite
            Timer = False
        End If
        If Brown(P).Height + Brown(P).Top < ZZ.Top Then
            Brown(1 - P).BackColor = vbWhite
            Timer = False
        End If
    End If
    If Brown(P).Height + Brown(P).Top > Me.ScaleHeight Then
        Brown(P).BackColor = vbWhite
        Timer = False
    End If
    If Blue(P).Height + Blue(P).Top > Me.ScaleHeight Then
        Blue(P).BackColor = vbWhite
        Timer = False
    End If
    For k = 0 To HisTop - 1
        His(P, k) = His(P, k + 1)
    Next k
Next P
End Sub

Private Sub Timer1_Timer()
Cap = Val(Cap) - 1
If Val(Cap) = 0 Then
    Timer = True
    Timer1 = False
    Cap.Visible = False
End If
End Sub

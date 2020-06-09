VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MOD"
   ClientHeight    =   6756
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11556
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   42
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6756
   ScaleWidth      =   11556
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Ans 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   3312
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   4200
      Width           =   4932
   End
   Begin VB.Label Board 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score"
      ForeColor       =   &H00FFFFFF&
      Height          =   888
      Left            =   4704
      TabIndex        =   4
      Top             =   0
      Width           =   2160
   End
   Begin VB.Label Midd 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mod"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   66
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   4776
      TabIndex        =   2
      Top             =   1680
      Width           =   2004
   End
   Begin VB.Label Num2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   66
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   7104
      TabIndex        =   1
      Top             =   1920
      Width           =   2640
   End
   Begin VB.Label Num1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   66
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   1824
      TabIndex        =   0
      Top             =   1920
      Width           =   2640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim Score As Long
Dim HighScore As Long

Const MyName As String = "How 2 Become A Calculater"

Private Sub Ans_KeyPress(KeyAscii As Integer)
With Ans
    Select Case KeyAscii
        Case 1  'Ctrl+A
            .SelStart = 0
            .SelLength = Len(.Text)
        Case 13 'ENTER
            ENTER
        Case 27 'ESC
            .Text = ""
    End Select
End With
End Sub

Private Sub ENTER()
If Val(Ans) = Val(Num1) Mod Val(Num2) Then
    Score = Score + 4
    NXT
Else
    Beep
    Ans = ""
End If
End Sub

Private Sub NXT()
Ans = ""
Dim Standard As Long
Standard = (Score / 19) ^ 2
If Rnd < 0.19 Then
    Num1 = Int(Rnd * Standard) + 3
Else
    Num1 = Standard - Int(Rnd * Standard * 0.19) + 3
End If
Num2 = Int(Num1 * (Rnd / 2))
Num2 = Val(Num2) + Int(Rnd * 4) - 1
If Val(Num2) > Val(Num1) Then
    Num2 = Num1
Else
    If Val(Num2) <= 1 Then
        Num2 = 2
    End If
End If
Board = Score
End Sub

Private Sub Form_Load()
Randomize
NXT
End Sub


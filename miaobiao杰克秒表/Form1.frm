VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "杰克秒表"
   ClientHeight    =   3204
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8112
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3204
   ScaleWidth      =   8112
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Main 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   1440
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   99
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1980
      Index           =   1
      Left            =   3045
      TabIndex        =   2
      Top             =   612
      Width           =   2028
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   99
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1980
      Index           =   0
      Left            =   579
      TabIndex        =   1
      Top             =   612
      Width           =   2028
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   99
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1980
      Index           =   2
      Left            =   5505
      TabIndex        =   0
      Top             =   612
      Width           =   2028
   End
   Begin VB.Shape ZZ 
      BorderColor     =   &H8000000F&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   612
      Left            =   0
      Top             =   0
      Width           =   612
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BG As Long
Const BC As Long = &HCC0000
Dim OldS As Integer
Dim Ms As Integer
Dim Delt As Integer
Const fR As Integer = 6

Private Sub Form_Click()
Main = Not Main
OldS = Second(Now)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 32 Then
    Form_Click
End If
End Sub

Private Sub Form_Load()
Delt = 60
BG = Me.BackColor
Main.Interval = 1
ZZ.Height = Me.Height
ZZ.Width = 0
End Sub

Private Sub L_Click(Index As Integer)
Form_Click
End Sub

Private Sub Main_Timer()
Dim S As Integer, M As Integer, H As Integer
If Second(Now) <> OldS Then
    OldS = Second(Now)
    H = Val(L(0))
    M = Val(L(1))
    S = Val(L(2))
    If S = 59 Then
        If M = 59 Then
            H = H + 1
            M = 0
        Else
            M = M + 1
        End If
        S = 0
    Else
        S = S + 1
    End If
    L(0) = Format(H, "00")
    L(1) = Format(M, "00")
    L(2) = Format(S, "00")
    If Me.BackColor = BG Then
        Me.BackColor = BC
        ZZ.FillColor = BG
    Else
        Me.BackColor = BG
        ZZ.FillColor = BC
    End If
    ZZ.BorderColor = ZZ.FillColor
    Delt = Delt + Ms * 5
    Delt = Delt / 6
    Ms = 0
    Debug.Print Delt
    ZZ.Width = 0
Else
    Ms = Ms + 1
    If Ms Mod fR = 0 Then
        ZZ.Width = ZZ.Width + Me.Width / Delt * fR
    End If
End If
End Sub

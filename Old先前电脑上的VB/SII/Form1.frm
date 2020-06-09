VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stanley博士的家 II 光之石获取练习"
   ClientHeight    =   6468
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10248
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6468
   ScaleWidth      =   10248
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4680
      Top             =   3000
   End
   Begin VB.HScrollBar Bs 
      Height          =   852
      Left            =   7680
      Max             =   4
      TabIndex        =   4
      Top             =   3480
      Value           =   4
      Width           =   2172
   End
   Begin VB.HScrollBar Gs 
      Height          =   852
      Left            =   7680
      Max             =   4
      TabIndex        =   2
      Top             =   2280
      Value           =   4
      Width           =   2172
   End
   Begin VB.HScrollBar Rs 
      Height          =   852
      Left            =   7680
      Max             =   4
      TabIndex        =   1
      Top             =   1080
      Value           =   4
      Width           =   2172
   End
   Begin VB.Shape S2 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   612
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   492
   End
   Begin VB.Line Bl 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      X1              =   3120
      X2              =   4680
      Y1              =   3240
      Y2              =   1440
   End
   Begin VB.Line Rl 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   3120
      X2              =   1440
      Y1              =   3240
      Y2              =   1440
   End
   Begin VB.Line Gl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   5
      X1              =   3120
      X2              =   3120
      Y1              =   960
      Y2              =   3240
   End
   Begin VB.Label B 
      BackColor       =   &H80000008&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   972
      Left            =   6960
      TabIndex        =   5
      Top             =   3480
      Width           =   612
   End
   Begin VB.Label G 
      BackColor       =   &H80000008&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   972
      Left            =   6960
      TabIndex        =   3
      Top             =   2280
      Width           =   612
   End
   Begin VB.Label R 
      BackColor       =   &H80000008&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   972
      Left            =   6960
      TabIndex        =   0
      Top             =   1080
      Width           =   612
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   612
      Left            =   2800
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   612
   End
   Begin VB.Line Tr 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   15
      X1              =   3120
      X2              =   3120
      Y1              =   5040
      Y2              =   6120
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   10
      X1              =   3120
      X2              =   3120
      Y1              =   3240
      Y2              =   4440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yLw As Boolean
Dim trR As Byte, trG As Byte, trB As Byte
Private Sub Form_Load()
Randomize
trR = Int(Rnd() * 5)
trG = Int(Rnd() * 5)
trB = Int(Rnd() * 5)
Tr.BorderColor = RGB(trR * 63, trG * 63, trB * 63)
Call rNw
End Sub

Private Sub rNw()
L.BorderColor = RGB(Rs.Value * 63, Gs.Value * 63, Bs.Value * 63)
Rl.BorderColor = RGB(Rs.Value * 63, 0, 0)
Gl.BorderColor = RGB(0, Gs.Value * 63, 0)
Bl.BorderColor = RGB(0, 0, Bs.Value * 63)
If Rs.Value = trR And Gs.Value = trG And Bs.Value = trB Then
    MsgBox "成功！", vbOKOnly + vbInformation, "恭喜"
    Call Form_Load
End If
End Sub

Private Sub Rs_Change()
Call rNw
End Sub

Private Sub Gs_Change()
Call rNw
End Sub

Private Sub Bs_Change()
Call rNw
End Sub

Private Sub Timer1_Timer()
If yLw Then
    S2.FillColor = RGB(255, 255, 0)
    yLw = False
Else
    S2.FillColor = 0
    yLw = True
End If
End Sub

VERSION 5.00
Begin VB.Form Loader 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   LinkTopic       =   "Form2"
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer TpGX 
      Interval        =   5000
      Left            =   5160
      Top             =   3600
   End
   Begin VB.Timer Loading 
      Interval        =   50
      Left            =   4320
      Top             =   3360
   End
   Begin VB.Label Tips 
      BackColor       =   &H00000000&
      Caption         =   "温馨提醒"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2412
      Left            =   960
      TabIndex        =   0
      Top             =   3240
      Width           =   7452
   End
   Begin VB.Shape ZheBi 
      FillStyle       =   0  'Solid
      Height          =   852
      Left            =   480
      Top             =   6240
      Width           =   8532
   End
   Begin VB.Image Logo 
      Height          =   7200
      Left            =   0
      Picture         =   "Form2.frx":0000
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Randomize
Call TpGX_Timer
With Logo
    .Picture = LoadPicture(App.Path & "\Files\Logo.jpg")
    Me.Width = .Width
    Me.Height = .Height
End With
End Sub

Private Sub Loading_Timer()
With ZheBi
    .Width = .Width - 50
    .Left = .Left + 50
    If .Width < 50 Then
        Loading.Interval = 0
        TpGX.Interval = 0
        Tips.FontSize = 120
        Tips.Caption = " PLAY"
    End If
End With
End Sub

Private Sub TpGX_Timer()    '更新小贴士
Dim W As Byte
W = Int(Rnd() * 10)
Select Case W
    Case 1
        
End Select
End Sub

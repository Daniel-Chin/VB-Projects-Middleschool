VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5976
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7044
   LinkTopic       =   "Form1"
   ScaleHeight     =   5976
   ScaleWidth      =   7044
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label L 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   3000
      TabIndex        =   0
      Top             =   2760
      Width           =   972
   End
   Begin VB.Image C 
      Height          =   372
      Left            =   2880
      Top             =   3600
      Width           =   972
   End
   Begin VB.Image B 
      Height          =   372
      Left            =   4200
      Top             =   3120
      Width           =   972
   End
   Begin VB.Image A 
      Height          =   372
      Left            =   3000
      Top             =   2520
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim Tms As Byte
Const S As Integer = -200, sS As Integer = 36
Dim sSs()

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    For k = sSs(Tms) To sSs(Tms + 1) Step S
        B.Top = k
        DoEvents
        Sleep sS
    Next k
    Tms = Tms + 1
End If
End Sub

Private Sub Form_Load()
sSs = Array(5200, 3800, 2300, 800, -700, -2100)
Me.BackColor = 0
With A
    .Picture = LoadPicture("D:\WDY\Wechat\Bar.bmp")
    .Top = 500
    .Left = 500
    .ZOrder
End With
With B
    .Picture = LoadPicture("D:\WDY\Wechat\chat.bmp")
    .Top = 5200
    .Left = 500
    .ZOrder 1
End With
With L
    .Top = 500
    .Left = 500
    .Height = 99999
    .Width = 4900
    .ZOrder 1
End With
With C
    .Picture = LoadPicture("D:\WDY\Wechat\Bt.bmp")
    .Top = Me.Height - .Height - 500
    .Left = 500
End With
End Sub


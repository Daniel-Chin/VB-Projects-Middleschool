VERSION 5.00
Begin VB.Form Zhu 
   ClientHeight    =   4872
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6900
   Icon            =   "Zhu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4872
   ScaleWidth      =   6900
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Pad 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   25.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   972
   End
End
Attribute VB_Name = "Zhu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim k As Integer

Dim His As String

Private Sub Form_Load()
Randomize
Load AI
Caption = "贾维斯"
End Sub

Private Sub Form_Resize() '调整Pad大小
    With Pad
        .Width = Width - 266
        .Height = Height - 666
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Un
End Sub

Private Sub Pad_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 27 'ESC
        Un
    Case 10 'Ctrl+Enter
        AI.Speak Pad
        His = Pad
        Pad = ""
    Case 19 'Ctrl + S
        Pad = His
End Select
End Sub


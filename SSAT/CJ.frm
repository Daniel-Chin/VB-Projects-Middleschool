VERSION 5.00
Begin VB.Form CJ 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "成就"
   ClientHeight    =   6180
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   11340
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "<<返回"
      Height          =   372
      Left            =   0
      MouseIcon       =   "CJ.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   0
      Width           =   972
   End
   Begin VB.Label LB 
      Alignment       =   2  'Center
      Caption         =   "你很有成就感！！！"
      Height          =   1932
      Left            =   3120
      TabIndex        =   0
      Top             =   1320
      Width           =   3492
   End
End
Attribute VB_Name = "CJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
With LB
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height / 2
    .FontSize = 56
    .BackColor = vbWhite
    .Caption = Chr(10) & .Caption
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'#######
End Sub


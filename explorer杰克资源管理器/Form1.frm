VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   6852
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11124
   LinkTopic       =   "Form1"
   ScaleHeight     =   6852
   ScaleWidth      =   11124
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.FileListBox File1 
      Height          =   252
      Hidden          =   -1  'True
      Left            =   5880
      System          =   -1  'True
      TabIndex        =   1
      Top             =   3360
      Width           =   972
   End
   Begin VB.DirListBox Dir1 
      Height          =   276
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Pth As String

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        Dir1.Path = Dir1.List(Dir1.ListIndex)
    Case 8
        
End Select
End Sub

Private Sub Form_Load()
Pth = Command
If Len(Pth) <= 1 Then
    Pth = "D:\"
End If
Dir1.FontSize = 33
File1.FontSize = Dir1.FontSize
End Sub

Private Sub Form_Resize()
With Dir1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width / 2
End With
With File1
    .Top = 0
    .Left = Me.Width / 2
    .Width = .Left
    .Height = Me.Height
End With
End Sub

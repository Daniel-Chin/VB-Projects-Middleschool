VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   166
      Left            =   1320
      Top             =   960
   End
   Begin VB.Image Image1 
      Height          =   12
      Left            =   1320
      Top             =   0
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ID As Integer

Private Sub Form_Load()
MsgBox "可能有文件将被删除：LXJP", vbCritical
End Sub

Private Sub Form_Resize()
Image1.Width = Width
Image1.Height = Me.Height
End Sub

Private Sub Timer1_Timer()
If Clipboard.GetData <> Image1.Picture Then
    Image1.Picture = Clipboard.GetData
    SavePicture Clipboard.GetData, "D:\LXJP\" & ID & ".jpg"
    ID = ID + 1
End If
End Sub

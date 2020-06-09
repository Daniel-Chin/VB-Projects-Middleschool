VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6804
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9744
   LinkTopic       =   "Form1"
   ScaleHeight     =   6804
   ScaleWidth      =   9744
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Image Img 
      Height          =   372
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Img.Picture = LoadPicture(Command)
End Sub

Private Sub Form_Resize()
Img.Move 0, 0, Me.Width, Me.Height
End Sub

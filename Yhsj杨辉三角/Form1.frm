VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6312
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   8796
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   526
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   733
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Pad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4572
      Left            =   0
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   371
      TabIndex        =   0
      Top             =   0
      Width           =   4452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Size As Long = 1024
Const MooD As Integer = 2
Dim DBD As Integer

Private Sub Form_Load()
Show
Pad.Move 0, 0, Size, Size
DBD = Int(255 / MooD)
Dim X As Long, Y As Long
Dim Slot As Integer, GB As Integer
Dim L(1 To Size) As Integer
L(1) = 1
Draw 1, 1, 1
For X = 2 To Size
    For Y = 1 To X
        If Y = 1 Then
            Draw X, Y, 1
            Slot = 1
        Else
            If Y = X Then
                Draw X, Y, 1
                L(Y) = 1
                L(Y - 1) = Slot
            Else
                GB = L(Y) + L(Y - 1)
                GB = GB Mod MooD
                Draw X, Y, GB
                L(Y - 1) = Slot
                Slot = GB
            End If
        End If
    Next Y
Next X
End Sub

Private Sub Draw(X As Long, Y As Long, Color As Integer)
Pad.PSet (X - Y + 1, Y), RGB(Color * DBD, Color * DBD, Color * DBD)
End Sub

Private Sub Pad_Click()
SavePicture Pad.Image, "d:\1.bmp"
MsgBox "Saved"
End Sub

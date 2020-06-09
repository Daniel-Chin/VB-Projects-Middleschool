VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   3744
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SC As Long
Dim Min As Double

Private Sub Command1_Click()
SC = SC * 2
Cls
Print SC
Min = 1
End Sub

Private Sub Form_Load()
Dim R As Double, N As Long, Cha As Double
Min = 1
Show
SC = 8
Randomize
Do
R = Rnd
N = Int(Rnd * SC) + 1
Cha = B(R, N) - A(R, N)
If Cha < Min Then
    Min = Cha
    Cls
    Print Min
End If
If Cha < 0 Then Stop
DoEvents
Loop
End Sub

Private Function A(R As Double, N As Long)
A = 2 * R ^ N - R ^ (2 * N)
End Function
Private Function B(R As Double, N As Long)
B = (2 * R - R ^ 2) ^ N
End Function

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
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   1080
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim CMD As String
CMD = Command
If InStr(Command, Chr(34)) = 1 Then
    CMD = Mid(CMD, 2, 999)
    CMD = Left(CMD, InStr(CMD, Chr(34)) - 1)
End If
Open CMD For Input As #1
Dim S As String
Do Until EOF(1)
Line Input #1, S
Text1 = Text1 & Chr(10) & Chr(13) & Chr(10) & S
Loop
End Sub

Private Sub Form_Resize()
Text1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

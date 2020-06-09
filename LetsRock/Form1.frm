VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4644
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8676
   LinkTopic       =   "Form1"
   ScaleHeight     =   4644
   ScaleWidth      =   8676
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   2160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Size As Single = 0.66
Dim Bin As String

Private Sub Form_Load()
Randomize
Show
Me.FontSize = 36
Print "I am control. "
Me.Caption = "Initializing"
End Sub

Private Sub Timer1_Timer()
Bin = App.Path & "\bin\"
Dim File(1 To 99999) As String
Dim k As Integer
k = 1
Dim TsT As String
TsT = Dir(Bin)
Do Until TsT = ""
    File(k) = TsT
    k = k + 1
    TsT = Dir
Loop
Dim fTop As Integer
fTop = k - 1
Me.Caption = fTop
Dim GL As Single
GL = 1 - Size ^ (fTop)
For k = 1 To fTop
    If Rnd() < GL Then
        Kill Bin & File(k)
        File(k) = ""
    End If
Next k
For k = 1 To fTop
    If File(k) <> "" Then
        Shell Bin & File(k), vbNormalFocus
    End If
Next k
End Sub

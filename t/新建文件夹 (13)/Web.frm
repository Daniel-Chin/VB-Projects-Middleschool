VERSION 5.00
Begin VB.Form Web 
   Caption         =   "股市模拟服务端"
   ClientHeight    =   5412
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9852
   LinkTopic       =   "Form1"
   ScaleHeight     =   5412
   ScaleWidth      =   9852
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Prompt 
      Height          =   4572
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   4092
   End
   Begin VB.Timer Timer 
      Interval        =   19
      Left            =   4440
      Top             =   2520
   End
End
Attribute VB_Name = "Web"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim IntPerS As Integer
Dim IntLastS As Integer
Dim LastS As Integer

Private Sub Form_Load()
Market.Show
End Sub

Private Sub Timer_Timer()
IntLastS = IntLastS + 1
If Second(Now) <> LastS Then
    LastS = Second(Now)
    IntPerS = IntLastS
    IntLastS = 0
    'pop IntPerS
End If
End Sub

Public Function NowMs() As Integer
NowMs = IntLastS / IntPerS * 1000
Debug.Assert IntLastS <= IntPerS
End Function

Private Sub Pop(MSG As String)
Prompt = Prompt & MSG & vbCrLf
'Prompt.BackColor = &HFF0000 - Prompt.BackColor
Prompt.SelStart = Len(Prompt)
Debug.Print MSG
End Sub

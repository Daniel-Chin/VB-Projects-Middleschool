VERSION 5.00
Begin VB.Form Market 
   Caption         =   "Market"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   3744
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   972
   End
   Begin VB.Timer Debuffer 
      Interval        =   19
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Timer BankShrinker 
      Interval        =   1000
      Left            =   1560
      Top             =   720
   End
End
Attribute VB_Name = "Market"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typBuff
    S As Integer
    MS As Integer
End Type

Dim Buff(1 To 999) As typBuff
Dim BuffNum As Integer

Public Sub AddBuff(S As Integer, MS As Integer)
    BuffNum = BuffNum + 1
    With Buff(BuffNum)
        .MS = MS
        .S = S
    End With
End Sub

Private Sub Command1_Click()
AddBuff Second(Now), Web.NowMs
End Sub

Private Sub Debuffer_Timer()
Dim k As Integer
Dim Delta As Long
Dim AllClear As Boolean
Do Until AllClear
    AllClear = True
    For k = 1 To BuffNum
        With Buff(k)
            Delta = (Second(Now) - .S) * 1000 + Web.NowMs - .MS
            Delta = (Delta + 60000) Mod 60000
            If Delta > 1000 Then
                Debug.Print "有问题！"
                Debug.Print .S; .MS; Second(Now); Web.NowMs
            End If
            If Delta >= 500 Then
                Me.BackColor = Rnd * 255
                .MS = Buff(BuffNum).MS
                .S = Buff(BuffNum).S
                BuffNum = BuffNum - 1
                AllClear = False
            End If
        End With
    Next k
Loop
End Sub

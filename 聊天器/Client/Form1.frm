VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5976
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   10176
   LinkTopic       =   "Form1"
   ScaleHeight     =   5976
   ScaleWidth      =   10176
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Pop 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   -120
      Visible         =   0   'False
      Width           =   4932
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Web.Send InputBox("输入消息")
End Sub

Private Sub Form_Load()
Web.Show
Web.Connect InputBox("IP:", , ""), 1875
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Web
End Sub

Public Sub InMSG(strData As String)
    Dim S() As String
    S = Split(strData, "`")
    Dim L As Integer
    L = Pop.UBound + 1
    Load Pop(L)
    With Pop(L)
        .BackColor = Val(S(0))
        .Caption = S(1)
        .Left = Pop(L - 1).Left
        .Top = Pop(L - 1).Top + Pop(L - 1).Height + Pop(0).Height * 0.3
        If .Top + .Height > Me.ScaleHeight Then
            .Top = Pop(1).Top
            .Left = Pop(0).Left + Pop(0).Width
        End If
        .AutoSize = True
        .Visible = True
    End With
End Sub

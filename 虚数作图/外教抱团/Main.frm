VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "bG"
   ClientHeight    =   2316
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Roller 
      Interval        =   10000
      Left            =   2040
      Top             =   600
   End
   Begin VB.Timer Starter 
      Interval        =   2666
      Left            =   2040
      Top             =   1200
   End
   Begin VB.Timer BS 
      Interval        =   128
      Left            =   1320
      Top             =   960
   End
   Begin VB.Label bT 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   972
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowCursor Lib "user32" (ByVal BShow As Long) As Long

Dim YsCase As Byte, MainR As Integer, MainB As Integer, MainG As Integer
Dim BsXs As Integer

Dim Tms As Long

Private Sub BS_Timer()
Tms = Tms + 1
If Tms = 1164 Then
    Tms = 0
    Shell App.Path & "\JMPlayer.exe " & App.Path & "\Stanley.mp3"
End If
Select Case YsCase
    Case 0
        MainR = MainR + BsXs
        If MainR >= 255 Then YsCase = 1: MainR = 255 'R
    Case 1
        MainG = MainG + BsXs
        If MainG >= 255 Then YsCase = 2: MainG = 255 'RG
    Case 2
        MainR = MainR - BsXs
        If MainR <= 0 Then YsCase = 3: MainR = 0 'G
    Case 3
        MainB = MainB + BsXs
        If MainB >= 255 Then YsCase = 4: MainB = 255 'GB
    Case 4
        MainG = MainG - BsXs
        If MainG <= 0 Then YsCase = 5: MainG = 0 'B
    Case 5
        MainR = MainR + BsXs
        If MainR >= 255 Then YsCase = 6: MainR = 255 'BR
    Case 6
        MainB = MainB - BsXs
        If MainB <= 0 Then YsCase = 1: MainB = 0 'R
End Select
Me.BackColor = RGB(MainR, MainG, MainB)
bT.BackColor = vbWhite - Me.BackColor
bT.ForeColor = Me.BackColor
End Sub

Public Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 27
        ShowCursor True
        End
    Case 119
        bT.Height = bT.Height - 6
    Case 115
        bT.Height = bT.Height + 6
End Select
End Sub

Private Sub Form_Load()
Tms = 281
BsXs = 4
With Screen
    Me.Move 0, 0, .Width, .Height
End With
With bT
    .Left = 0
    .Width = Me.Width
    .Height = 3666
    .Top = (Me.Height - .Height) / 2
    .FontSize = 66
    .Caption = "Group 1" & Chr(10) & "Presentation"
    .FontBold = True
End With
ShowCursor False
Roller_Timer
Shell App.Path & "\JMPlayer.exe " & App.Path & "\Start.mp3"
End Sub

Private Sub Roller_Timer()
If Tag = "a" Then
    BS.Interval = 16
    Tag = ""
    BsXs = 16
    Roller.Interval = 4000
Else
    BS.Interval = 66
    Tag = "a"
    BsXs = 4
    Roller.Interval = 20000
End If
End Sub

Private Sub Starter_Timer()
Starter.Tag = Val(Starter.Tag) + 1
Select Case Starter.Tag
    Case 1
        Load F1
    Case 2
        Load F2
    Case 3
        Load F3
    Case 4
        Starter = False
End Select
End Sub

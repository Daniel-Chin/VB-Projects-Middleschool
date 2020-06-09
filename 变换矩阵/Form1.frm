VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6696
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11664
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   ScaleHeight     =   6696
   ScaleWidth      =   11664
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Top             =   924
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   924
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Pixel
    Left As Long
    Top As Long
End Type

Private Type Point
    X As Single
    Y As Single
End Type

Private Type Line
    X1 As Single
    Y1 As Single
    X2 As Single
    y2 As Single
    Color As Long
    Connected As Boolean
    D1 As Integer
    D2 As Integer
End Type

Dim O As Pixel
Dim Bily As Double  '比例尺

Dim Xian(1 To 999) As Line
Dim XianTop As Integer

Dim Dot(0 To 999) As Point
Dim DotTop As Integer

Dim DotOnlyMode As Boolean

Dim DefaultColor As Long

Dim aa As Single
Dim bb As Single
Dim cc As Single
Dim dd As Single
Const Size As Integer = 13

Private Sub Form_Load()
With Me
    .FontSize = 30
    .Caption = "杰克坐标系"
    DefaultColor = .ForeColor
End With
Bily = 466

'''''''''''''''''''''''''
'自定义区域
DotNew -1, -2
DotNew 0, 3
DotNew 2, 5
DotConnect 1, 2
DotConnect 1, 3
DotConnect 2, 3
DotNew 1, 1
DotNew 1, 2
DotNew 2, 2
DotNew 2, 1
DotConnect 4, 5
DotConnect 5, 6
DotConnect 6, 7
DotConnect 7, 4
DotNew 1, 3
DotNew 2, 3
DotConnect 8, 9
DotNew -88, 0
DotNew 88, 0
DotConnect 10, 11
DotNew 0, -88
DotNew 0, 88
DotConnect 12, 13

DotNew -1, -2
DotNew 0, 3
DotNew 2, 5
eDotConnect 1, 2
eDotConnect 1, 3
eDotConnect 2, 3
DotNew 1, 1
DotNew 1, 2
DotNew 2, 2
DotNew 2, 1
eDotConnect 4, 5
eDotConnect 5, 6
eDotConnect 6, 7
eDotConnect 7, 4
DotNew 1, 3
DotNew 2, 3
eDotConnect 8, 9
DotNew -8, 0
DotNew 8, 0
eDotConnect 10, 11
DotNew 0, -8
DotNew 0, 8
eDotConnect 12, 13

aa = 1
bb = 0
cc = 0
dd = 1
Text1(0) = aa
Text1(1) = bb
Text1(2) = cc
Text1(3) = dd
Dim k As Integer
For k = 0 To 3
    Text1(k).FontSize = 20
Next k
'''''''''''''''''''''''''
End Sub

Private Sub Form_Resize()
O.Left = Me.Width / 2
O.Top = Me.Height / 2
Paint
End Sub

Private Sub Paint()
Dim k As Integer
''''''''''''''''''''
For k = 1 To Size
    With Dot(k)
        Dot(k + Size).X = .X * aa + .Y * bb
        Dot(k + Size).Y = .X * cc + .Y * dd
    End With
Next k
''''''''''''''''''''
For k = 1 To XianTop
    With Xian(k)
        If .Connected = True Then
            .X1 = Dot(.D1).X
            .Y1 = Dot(.D1).Y
            .X2 = Dot(.D2).X
            .y2 = Dot(.D2).Y
        End If
    End With
Next k
Cls
Me.ForeColor = DefaultColor
Line (0, O.Top)-(Me.Width, O.Top)
Line (O.Left, 0)-(O.Left, Me.Height)
If XianTop >= 1 Then
    For k = 1 To XianTop
        With Xian(k)
            Me.ForeColor = .Color
            Line (D2P(.X1), D2P(.Y1, "y"))-(D2P(.X2), D2P(.y2, "y"))
        End With
    Next k
End If
End Sub

Private Sub XianNew(X1 As Single, Y1 As Single, X2 As Single, y2 As Single, Optional Color As Long = -1, Optional DoPaint As Boolean = True)
XianTop = XianTop + 1
If Color = -1 Then
    Color = Me.ForeColor
End If
With Xian(XianTop)
    .X1 = X1
    .X2 = X2
    .Y1 = Y1
    .y2 = y2
    .Color = Color
End With
If DoPaint Then
    Paint
End If
End Sub

Private Sub DotNew(X As Single, Y As Single)
DotTop = DotTop + 1
With Dot(DotTop)
    .X = X
    .Y = Y
End With
If DotOnlyMode Then
    DotConnect DotTop, 0
End If
Paint
End Sub

Private Function D2P(D As Single, Optional Axis As String = "x") As Long  'Dot2Pixel
D2P = D * Bily
If Axis = "x" Then
    D2P = D2P + O.Left
Else
    D2P = -D2P + O.Top
End If
End Function

Private Sub DotConnect(D1 As Integer, D2 As Integer, Optional Color As Long = -1, Optional DoPaint As Boolean = True)
XianNew Dot(D1).X, Dot(D1).Y, Dot(D2).X, Dot(D2).Y, Color, DoPaint
With Xian(XianTop)
    .Connected = True
    .D1 = D1
    .D2 = D2
End With
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
    Case 0
        aa = Val(Text1(Index))
    Case 1
        bb = Val(Text1(Index))
    Case 2
        cc = Val(Text1(Index))
    Case 3
        dd = Val(Text1(Index))
End Select
Paint
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 23 Then
    Text1(Index) = Val(Text1(Index)) + 0.1
End If
If KeyAscii = 19 Then
    Text1(Index) = Val(Text1(Index)) - 0.1
End If
End Sub

Private Sub eDotConnect(D1 As Integer, D2 As Integer, Optional Color As Long = &HFFFF&, Optional DoPaint As Boolean = True)
D1 = D1 + Size
D2 = D2 + Size
XianNew Dot(D1).X, Dot(D1).Y, Dot(D2).X, Dot(D2).Y, Color, DoPaint
With Xian(XianTop)
    .Connected = True
    .D1 = D1
    .D2 = D2
End With
End Sub


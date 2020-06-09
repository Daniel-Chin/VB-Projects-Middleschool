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
   StartUpPosition =   3  '窗口缺省
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
End Type

Dim O As Pixel
Dim Bily As Double  '比例尺

Dim Xian(1 To 999) As Line
Dim XianTop As Integer

Dim Dot(0 To 999) As Point
Dim DotTop As Integer

Dim DotOnlyMode As Boolean

Dim DefaultColor As Long

Private Sub Form_Load()
With Me
    .FontSize = 30
    .Caption = "杰克坐标系"
    DefaultColor = .ForeColor
End With
Bily = 466

'''''''''''''''''''''''''
'自定义区域
'''''''''''''''''''''''''
End Sub

Private Sub Form_Resize()
O.Left = Me.Width / 2
O.Top = Me.Height / 2
Paint
End Sub

Private Sub Paint()
Cls
Me.ForeColor = DefaultColor
Line (0, O.Top)-(Me.Width, O.Top)
Line (O.Left, 0)-(O.Left, Me.Height)
Dim k As Integer
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
    DotConnect Dot(DotTop), Dot(0)
End If
End Sub

Private Function D2P(D As Single, Optional Axis As String = "x") As Long  'Dot2Pixel
D2P = D * Bily
If Axis = "x" Then
    D2P = D2P + O.Left
Else
    D2P = -D2P + O.Top
End If
End Function

Private Sub DotConnect(D1 As Point, D2 As Point, Optional Color As Long = -1, Optional DoPaint As Boolean = True)
XianNew D1.X, D1.Y, D2.X, D2.Y, Color, DoPaint
End Sub

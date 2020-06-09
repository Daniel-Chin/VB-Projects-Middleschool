VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "º¯ÊýÍ¼Ïñ"
   ClientHeight    =   6276
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   11352
   LinkTopic       =   "Form1"
   ScaleHeight     =   6276
   ScaleWidth      =   11352
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xMIN As Single
Dim xMAX As Single
Dim yMIN As Single
Dim yMAX As Single
Dim DoStretch As Boolean '¾ÍºöÂÔyMax
Const GridStep As Single = 1

Private Sub Form_Load()
xMIN = -4
xMAX = 4
yMIN = -1
yMAX = 5
DoStretch = False
End Sub

Private Sub Graph(Keep As Boolean)
If Not Keep Then
    Cls
End If
Dim X As Single
Dim NowX As Long, NowY As Long
Dim LastX As Long, LastY As Long
LastX = L(xMIN)
LastY = T(f(xMIN))
For X = xMIN To xMAX Step (xMAX - xMIN) / 1024
    NowX = L(X)
    NowY = T(f(X))
    Line (NowX, NowY)-(LastX, LastY), vbYellow
    LastX = NowX
    LastY = NowY
Next X
'Line (0, T(yMAX))-(Me.ScaleWidth, T(yMAX))
End Sub

Private Function f(X As Single) As Single
f = X ^ 2
End Function

Private Function L(ZuoBiao As Single) As Long
L = Me.ScaleWidth * (ZuoBiao - xMIN) / (xMAX - xMIN)
End Function

Private Function T(ZuoBiao As Single) As Long
Dim YM As Single
If DoStretch Then
    YM = yMAX
Else
    YM = yMIN + (xMAX - xMIN) * Me.ScaleHeight / Me.ScaleWidth
End If
T = Me.ScaleHeight * (1 - (ZuoBiao - yMIN) / (YM - yMIN))
End Function

Private Sub Form_Resize()
Cls
Grid
Graph True
End Sub

Private Sub Grid()
Dim X As Single, Y As Single
For X = Int(xMIN / GridStep) * GridStep To xMAX Step GridStep
    Line (L(X), 0)-(L(X), Me.ScaleHeight)
Next X
For Y = Int(yMIN / GridStep) * GridStep To yMAX Step GridStep
    Line (0, T(Y))-(Me.ScaleWidth, T(Y))
Next Y
End Sub

VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6024
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   11268
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   ScaleHeight     =   6024
   ScaleWidth      =   11268
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4932
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   360
      Width           =   10212
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const tp As Long = 150000000
Dim S(2 To tp) As Boolean

Private Sub Form_Load()
Show
Text1.Visible = False
Dim k As Long, f As Long, st As Long, l As Long, oldy As Long
Dim bL As Single, y As Long, ooy As Long
Dim CS As Double
bL = 7800
CS = 0.5
st = 3000000
l = 1
For k = 2 To tp
    If S(k) Then
        l = l + 1
    Else
        f = k
        Do Until f > tp / k
            f = f * k
            S(f) = True
        Loop
    End If
    If k Mod st = 0 Then
        Text1 = l & Chr(13) & Chr(10) & Int(k / tp * 100) & "%"
        y = l / (200) * Me.ScaleHeight
        Line (Me.ScaleWidth * (k - st) / tp, 0)-(Me.ScaleWidth * k / tp, bL / k ^ CS * Me.ScaleHeight), vbRed
        Line (Me.ScaleWidth * (k - st) / tp, oldy - ooy)-(Me.ScaleWidth * k / tp, y - oldy)
        DoEvents
        ooy = oldy
        oldy = y
    End If
Next k
k = k - 1
Me.Caption = l
Text1.Visible = False
'lim
End Sub

Private Sub Pr()
Text1 = ""
Dim k As Long
For k = 2 To tp
    Text1 = Text1 & Format(Abs(S(k)), " 0")
Next k
End Sub

Private Sub Tr()
Text1 = ""
Dim k As Long
For k = 2 To tp
    If S(k) Then
        Text1 = Text1 & k & " "
    End If
Next k
End Sub

Private Sub lim()
Dim l As Double, k As Long, st As Long, oldl As Double
Text1.Visible = True
st = 10000
l = 1 '''''
For k = 2 To tp
    If Not S(k) Then
        l = l + 1 / (1 - k)
    End If
    If k Mod st = 0 Then
        Text1 = l & Chr(13) & Chr(10) & Int(k / tp * 100) & "%"
        Line (Me.ScaleWidth * (k - st) / tp, oldl / (-17) * Me.ScaleHeight)-(Me.ScaleWidth * k / tp, l / (-17) * Me.ScaleHeight)
        DoEvents
        oldl = l
    End If
Next k
Text1.Visible = False
Me.Caption = l
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

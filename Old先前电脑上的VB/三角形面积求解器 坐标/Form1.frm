VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   4128
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8964
   LinkTopic       =   "Form1"
   ScaleHeight     =   4128
   ScaleWidth      =   8964
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text6 
      Height          =   372
      Left            =   7332
      TabIndex        =   6
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox Text5 
      Height          =   372
      Left            =   7332
      TabIndex        =   5
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox Text4 
      Height          =   372
      Left            =   7332
      TabIndex        =   4
      Top             =   1200
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   732
      Left            =   6204
      TabIndex        =   3
      Top             =   3000
      Width           =   2052
   End
   Begin VB.TextBox Text3 
      Height          =   372
      Left            =   6156
      TabIndex        =   2
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   6156
      TabIndex        =   1
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   6156
      TabIndex        =   0
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   ","
      Height          =   252
      Left            =   7200
      TabIndex        =   10
      Top             =   2280
      Width           =   132
   End
   Begin VB.Label Label3 
      Caption         =   ","
      Height          =   252
      Left            =   7200
      TabIndex        =   9
      Top             =   1800
      Width           =   132
   End
   Begin VB.Label Label2 
      Caption         =   ","
      Height          =   252
      Left            =   7200
      TabIndex        =   8
      Top             =   1320
      Width           =   132
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   1800
      Y1              =   960
      Y2              =   1080
   End
   Begin VB.Line Line2 
      X1              =   600
      X2              =   1080
      Y1              =   1680
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   480
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   72
      Left            =   1740
      TabIndex        =   7
      Top             =   1740
      Width           =   72
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim v As Single, w As Single, a As Single, d As Single, s As Single, c As Single, Z As Single, K As Single, B As Single, Cent As Integer
    v = Val(Text1)
    w = Val(Text4)
    a = Val(Text2)
    d = Val(Text5)
    s = Val(Text3)
    c = Val(Text6)
''''````````'
        Cent = Label1.Top + Label1.Height / 2
        Line1.X1 = Cent + v * 120
        Line1.Y1 = Cent - w * 120
        Line2.X1 = Cent + a * 120
        Line2.Y1 = Cent - d * 120
        Line3.X1 = Cent + s * 120
        Line3.Y1 = Cent - c * 120
        Line1.X2 = Line3.X1
        Line1.Y2 = Line3.Y1
        Line2.X2 = Line1.X1
        Line2.Y2 = Line1.Y1
        Line3.X2 = Line2.X1
        Line3.Y2 = Line2.Y1
''''````````'
    K = (w - d) / (v - a)
    B = d - (a * (w - d)) / (v - a)
    Z = (s * K + B - c) * (v - a) / 2
    If Z < 0 Then
        Z = -Z
    End If
    Cls
    Print "S=" & Z
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.SetFocus
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5.SetFocus
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text6.SetFocus
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
    Text1.SetFocus
End If
End Sub


VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   6564
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11028
   LinkTopic       =   "Form11"
   ScaleHeight     =   6564
   ScaleWidth      =   11028
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   5040
      Top             =   3120
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   732
      Index           =   3
      Left            =   7560
      Top             =   1080
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   372
      Index           =   2
      Left            =   7250
      Top             =   4400
      Width           =   2292
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1150
      Index           =   1
      Left            =   2640
      Top             =   4680
      Width           =   372
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   612
      Index           =   0
      Left            =   2040
      Top             =   1080
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   972
      Index           =   3
      Left            =   7440
      Picture         =   "Form11.frx":0000
      Top             =   960
      Width           =   948
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   2
      Left            =   7080
      Picture         =   "Form11.frx":12D0
      Top             =   4200
      Width           =   2604
   End
   Begin VB.Image Image1 
      Height          =   1380
      Index           =   1
      Left            =   2520
      Picture         =   "Form11.frx":2995
      Top             =   4560
      Width           =   636
   End
   Begin VB.Image Image1 
      Height          =   828
      Index           =   0
      Left            =   1920
      Picture         =   "Form11.frx":3D32
      Top             =   960
      Width           =   1224
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t(0 To 3) As Integer, l(0 To 3) As Integer
Private Sub Form_Load()
Form11.BackColor = RGB(43, 47, 104)
For i = 0 To 3
    Image1(i).Picture = LoadPicture(App.Path & "\img\" & i & ".jpg")
Next i
t(0) = 100
l(0) = 100
t(1) = -80
l(1) = 90
t(2) = -120
l(2) = -110
t(3) = 60
l(3) = -70
End Sub
Private Sub Timer1_Timer()    'main
For i = 0 To 3
    With Image1(i)
        .Top = .Top + t(i)
        .Left = .Left + l(i)
        If .Top < 0 Or .Top > Form11.Height - .Height - 600 Then
            t(i) = -t(i)
        End If
        If .Left < 0 Or .Left > Form11.Width - .Width - 300 Then
            l(i) = -l(i)
        End If
    End With
    With Shape1(i)
        .Top = Image1(i).Top + 120
        .Left = Image1(i).Left + 120
    End With
Next i
End Sub

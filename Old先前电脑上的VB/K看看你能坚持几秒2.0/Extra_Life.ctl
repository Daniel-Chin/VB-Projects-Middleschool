VERSION 5.00
Begin VB.UserControl ExtraLife 
   BackColor       =   &H00400000&
   BackStyle       =   0  'Í¸Ã÷
   ClientHeight    =   1152
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1188
   ScaleHeight     =   1024
   ScaleMode       =   0  'User
   ScaleWidth      =   1024
   Begin VB.Timer Hcq 
      Left            =   120
      Top             =   360
   End
   Begin VB.Timer Tm 
      Interval        =   50
      Left            =   840
      Top             =   0
   End
   Begin VB.Line L 
      BorderColor     =   &H00008000&
      BorderWidth     =   5
      Index           =   3
      X1              =   403.394
      X2              =   610.263
      Y1              =   853.333
      Y2              =   853.333
   End
   Begin VB.Line L 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   5
      Index           =   2
      X1              =   827.475
      X2              =   827.475
      Y1              =   400
      Y2              =   613.333
   End
   Begin VB.Line L 
      BorderColor     =   &H00008000&
      BorderWidth     =   5
      Index           =   1
      X1              =   403.394
      X2              =   610.263
      Y1              =   213.333
      Y2              =   213.333
   End
   Begin VB.Line L 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   5
      Index           =   0
      X1              =   217.212
      X2              =   217.212
      Y1              =   400
      Y2              =   624
   End
   Begin VB.Line L 
      BorderColor     =   &H0000C000&
      BorderWidth     =   15
      Index           =   4
      X1              =   506.828
      X2              =   506.828
      Y1              =   250.667
      Y2              =   762.667
   End
   Begin VB.Line L 
      BorderColor     =   &H0000C000&
      BorderWidth     =   15
      Index           =   5
      X1              =   289.616
      X2              =   734.384
      Y1              =   512
      Y2              =   512
   End
End
Attribute VB_Name = "ExtraLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim EAB As Boolean
Dim a As Byte, b As Boolean
Dim c As Byte, d As Boolean
Private Sub Hcq_Tiextralifer()
Hcq.Interval = 0
EAB = False
End Sub
Private Sub Tm_Tiextralifer()
a = a + 6 * (b + 0.5)
If a > 50 Or a = 0 Then
    b = Not b
End If
c = c + 10 * (d + 0.5)
If c > 50 Or c = 0 Then
    d = Not d
End If
For i = 0 To 3
    L(i).BorderColor = RGB(a, 150 + a, a)
Next i
For i = 4 To 5
    L(i).BorderColor = RGB(c, 150 + c, c)
Next i
End Sub
Private Sub UserControl_Resize()
If EAB = False Then
    EAB = True
    Hcq.Interval = 10
    L(0).X1 = ExtraLife.ScaleWidth / 5
    L(0).X2 = ExtraLife.ScaleWidth / 5
    L(0).Y1 = ExtraLife.ScaleWidth / 5 * 2
    L(0).Y2 = ExtraLife.ScaleWidth / 5 * 3 '
    L(1).X1 = ExtraLife.ScaleWidth / 5 * 2
    L(1).X2 = ExtraLife.ScaleWidth / 5 * 3
    L(1).Y1 = ExtraLife.ScaleWidth / 5
    L(1).Y2 = ExtraLife.ScaleWidth / 5 '
    L(2).X1 = ExtraLife.ScaleWidth / 5 * 4
    L(2).X2 = ExtraLife.ScaleWidth / 5 * 4
    L(2).Y1 = ExtraLife.ScaleWidth / 5 * 2
    L(2).Y2 = ExtraLife.ScaleWidth / 5 * 3 '
    L(3).X1 = ExtraLife.ScaleWidth / 5 * 2
    L(3).X2 = ExtraLife.ScaleWidth / 5 * 3
    L(3).Y1 = ExtraLife.ScaleWidth / 5 * 4
    L(3).Y2 = ExtraLife.ScaleWidth / 5 * 4 '
    L(4).X1 = ExtraLife.ScaleWidth / 2
    L(4).X2 = ExtraLife.ScaleWidth / 2
    L(4).Y1 = ExtraLife.ScaleWidth / 5
    L(4).Y2 = ExtraLife.ScaleWidth / 5 * 4 '
    L(4).X1 = ExtraLife.ScaleWidth / 5
    L(4).X2 = ExtraLife.ScaleWidth / 5 * 4
    L(4).Y1 = ExtraLife.ScaleWidth / 2
    L(4).Y2 = ExtraLife.ScaleWidth / 2 '
End If
End Sub

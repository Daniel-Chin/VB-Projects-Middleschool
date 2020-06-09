VERSION 5.00
Begin VB.UserControl LinXin 
   BackStyle       =   0  'Í¸Ã÷
   ClientHeight    =   1212
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   804
   ScaleHeight     =   1212
   ScaleWidth      =   804
   Begin VB.Timer Floater 
      Interval        =   40
      Left            =   0
      Top             =   600
   End
   Begin VB.Timer Spinner 
      Interval        =   20
      Left            =   360
      Top             =   120
   End
   Begin VB.Timer bS 
      Interval        =   50
      Left            =   240
      Top             =   360
   End
   Begin VB.Line YX 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      X1              =   800
      X2              =   400
      Y1              =   400
      Y2              =   800
   End
   Begin VB.Line ZX 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      X1              =   0
      X2              =   400
      Y1              =   400
      Y2              =   800
   End
   Begin VB.Line YS 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      X1              =   800
      X2              =   400
      Y1              =   400
      Y2              =   0
   End
   Begin VB.Line ZS 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      X1              =   0
      X2              =   400
      Y1              =   400
      Y2              =   0
   End
End
Attribute VB_Name = "LinXin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim H As Single
Dim cM As Byte
Dim cIta As Single

Private Sub bS_Timer()
Select Case cM
    Case 0  'R+G
        ZX.BorderColor = ZX.BorderColor + RGB(0, 5, 0)
        ZS.BorderColor = ZS.BorderColor + RGB(0, 5, 0)
        YX.BorderColor = YX.BorderColor + RGB(0, 5, 0)
        YS.BorderColor = YS.BorderColor + RGB(0, 5, 0)
        If ZX.BorderColor = RGB(200, 200, 0) Then
            cM = 1
        End If
    Case 1  'RG-R
        ZX.BorderColor = ZX.BorderColor - RGB(5, 0, 0)
        ZS.BorderColor = ZS.BorderColor - RGB(5, 0, 0)
        YX.BorderColor = YX.BorderColor - RGB(5, 0, 0)
        YS.BorderColor = YS.BorderColor - RGB(5, 0, 0)
        If ZX.BorderColor = RGB(0, 200, 0) Then
            cM = 2
        End If
    Case 2  'G+B
        ZX.BorderColor = ZX.BorderColor + RGB(0, 0, 5)
        ZS.BorderColor = ZS.BorderColor + RGB(0, 0, 5)
        YX.BorderColor = YX.BorderColor + RGB(0, 0, 5)
        YS.BorderColor = YS.BorderColor + RGB(0, 0, 5)
        If ZX.BorderColor = RGB(0, 200, 200) Then
            cM = 3
        End If
    Case 3  'GB-G
        ZX.BorderColor = ZX.BorderColor - RGB(0, 5, 0)
        ZS.BorderColor = ZS.BorderColor - RGB(0, 5, 0)
        YX.BorderColor = YX.BorderColor - RGB(0, 5, 0)
        YS.BorderColor = YS.BorderColor - RGB(0, 5, 0)
        If ZX.BorderColor = RGB(0, 0, 200) Then
            cM = 4
        End If
    Case 4  'B+R
        ZX.BorderColor = ZX.BorderColor + RGB(5, 0, 0)
        ZS.BorderColor = ZS.BorderColor + RGB(5, 0, 0)
        YX.BorderColor = YX.BorderColor + RGB(5, 0, 0)
        YS.BorderColor = YS.BorderColor + RGB(5, 0, 0)
        If ZX.BorderColor = RGB(200, 0, 200) Then
            cM = 5
        End If
    Case 5  'BR-B
        ZX.BorderColor = ZX.BorderColor - RGB(0, 0, 5)
        ZS.BorderColor = ZS.BorderColor - RGB(0, 0, 5)
        YX.BorderColor = YX.BorderColor - RGB(0, 0, 5)
        YS.BorderColor = YS.BorderColor - RGB(0, 0, 5)
        If ZX.BorderColor = RGB(200, 0, 0) Then
            cM = 0
        End If
End Select
End Sub

Private Sub Floater_Timer()
Dim y As Integer
H = H + 0.12
If H > 6.28 Then
    H = H - 6.2831853
End If
y = 100 * (Sin(H + 1))
YS.Y1 = y + 500
ZS.Y1 = YS.Y1
ZX.Y1 = YS.Y1
YX.Y1 = YS.Y1
YS.Y2 = y + 100
ZS.Y2 = YS.Y2
YX.Y2 = y + 900
ZX.Y2 = YX.Y2
End Sub

Private Sub Spinner_Timer()
cIta = cIta + 0.04
If cIta >= 6.28 Then
    cIta = cIta - 6.2831853
End If
YS.X1 = 400 + Cos(cIta) * 400
YX.X1 = YS.X1
ZS.X1 = 400 - Cos(cIta) * 400
ZX.X1 = ZS.X1
End Sub

Private Sub UserControl_Initialize()
ZX.BorderColor = RGB(200, 0, 0)
ZS.BorderColor = RGB(200, 0, 0)
YX.BorderColor = RGB(200, 0, 0)
YS.BorderColor = RGB(200, 0, 0)
End Sub

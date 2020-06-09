VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5208
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   434
   ScaleMode       =   0  'User
   ScaleWidth      =   325
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Ne 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   1200
      ScaleHeight     =   372
      ScaleWidth      =   972
      TabIndex        =   0
      Top             =   2400
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Pi As Double = 3.14159265358979

Private Sub Ne_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
'Ne.Height = 2828
'Ne.Width = 2000
'Ne.Top = 0
'Show
Ne.Height = 2828
Ne.Width = 2000
Ne.ScaleMode = 0
Ne.ScaleHeight = 2828
Ne.ScaleWidth = 2000
Dim Wid As Long, Hei As Long
Wid = Ne.ScaleWidth
Hei = Ne.ScaleHeight
Dim X, Y, R, K, PreX, PreY, KK, KkK
X = Wid / 2
Y = 300
R = Y - 80
Dim FL As Long
FL = 50
For KkK = 0 To 8
    KK = Sqr(KkK) / Sqr(8) * FL
    K = 0
    Ne.DrawWidth = Round(10 - Sqr(KK) / 7 * 9)
    'Ne.DrawWidth = 7
    Ne.ForeColor = CL(KK, FL)
    R = Y - 80 - KK
PreX = X + Cos(K) * R
PreY = Y + Sin(K) * R
    For K = 0 To 2 * Pi Step Pi / 1024
        Ne.Line (PreX, PreY)-(X + Cos(K) * R, Y + Sin(K) * R)
        PreX = X + Cos(K) * R
        PreY = Y + Sin(K) * R
    Next K
Next KkK
If 1 Then
'If 0 Then
    SavePicture Ne.Image, "D:\HB.bmp"
    Shell "explorer D:\HB.bmp", vbMaximizedFocus
'    Unload Me
End If
End Sub

Function CL(KK, FL) As Long
'0~FL
Dim D
D = Int(KK / FL * 256)
CL = RGB(D, D, D)
End Function

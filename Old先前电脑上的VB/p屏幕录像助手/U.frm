VERSION 5.00
Begin VB.Form U 
   BorderStyle     =   0  'None
   Caption         =   "ÆÁÄ»Â¼ÏñÖúÊÖ"
   ClientHeight    =   4548
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7992
   LinkTopic       =   "Form1"
   ScaleHeight     =   4548
   ScaleWidth      =   7992
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer rEnew 
      Interval        =   50
      Left            =   3480
      Top             =   2040
   End
End
Attribute VB_Name = "U"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim position As POINTAPI

Public hEi As Byte, wId As Byte
Public YanSe As String
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 27     'ESC
        End
    Case 100    'd
        wId = wId + 1
    Case 97     'a
        wId = wId - 1
    Case 119    'w
        hEi = hEi + 1
    Case 115    's
        hEi = hEi - 1
    Case 114    'r
        YanSe = "r"
    Case 103    'g
        YanSe = "g"
    Case 98     'b
        YanSe = "b"
End Select
End Sub

Private Sub Form_Load()
hEi = 2
wId = 7
YanSe = "g"
U.Height = 100
D.Height = 100
L.Width = 100
R.Width = 100
D.Visible = True
L.Visible = True
R.Visible = True
rEnew_Timer
End Sub

Public Sub rEnew_Timer()
Select Case YanSe
    Case "r"
        U.BackColor = vbRed
        D.BackColor = vbRed
        L.BackColor = vbRed
        R.BackColor = vbRed
    Case "g"
        U.BackColor = vbGreen
        D.BackColor = vbGreen
        L.BackColor = vbGreen
        R.BackColor = vbGreen
    Case "b"
        U.BackColor = vbBlue
        D.BackColor = vbBlue
        L.BackColor = vbBlue
        R.BackColor = vbBlue
End Select
GetCursorPos position
U.Top = (position.y - hEi * 10) * 12
U.Left = (position.x - wId * 10) * 12
U.Width = wId * 20 * 12
D.Top = (position.y + hEi * 10) * 12
D.Left = (position.x - wId * 10) * 12
D.Width = wId * 20 * 12
L.Top = (position.y - hEi * 10) * 12
L.Left = (position.x - wId * 10) * 12
L.Height = hEi * 20 * 12
R.Top = (position.y - hEi * 10) * 12
R.Left = (position.x + wId * 10) * 12
R.Height = hEi * 20 * 12
End Sub


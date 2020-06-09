VERSION 5.00
Begin VB.UserControl Jdt 
   BackColor       =   &H00000000&
   CanGetFocus     =   0   'False
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   Begin VB.Shape Light 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   492
      Index           =   0
      Left            =   1080
      Top             =   1320
      Width           =   492
   End
End
Attribute VB_Name = "Jdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Cl(0 To 1, 0 To 2) As Byte
Dim Bars As Byte
Dim m(0 To 255) As Boolean

Public Sub RenewFilledColor(R As Byte, G As Byte, B As Byte)
Cl(1, 0) = R
Cl(1, 1) = G
Cl(1, 2) = B
Renew
End Sub

Public Sub RenewEmptyColor(R As Byte, G As Byte, B As Byte)
Cl(0, 0) = R
Cl(0, 1) = G
Cl(0, 2) = B
Renew
End Sub

Public Sub RenewBackColor(R As Byte, G As Byte, B As Byte)
UserControl.BackColor = RGB(R, G, B)
End Sub

Public Sub RenewBars(ShuLiang As Byte)
ShuLiang = ShuLiang - 1
If ShuLiang <= 0 Then MsgBox "ERROR in 进度条 用户控件 更新灯数量", vbCritical + vbOKOnly, "ERROR": Stop
Dim k As Long
For k = 1 To Bars
    Unload Light(k)
Next k
'''''''''
Dim bh As Long, bw As Long, w As Long
w = UserControl.ScaleWidth
bw = Int(w / (ShuLiang + 1) / 5)
bh = Int(UserControl.ScaleHeight / 5)
'''''''''
For k = 0 To ShuLiang
    If k >= 1 Then Load Light(k)
    With Light(k)
        .Left = w / (ShuLiang + 1) * k + bw
        .Top = bh
        .Width = 3 * bw
        .Height = bh * 3
        .Visible = True
    End With
Next k
Bars = ShuLiang
End Sub

Public Sub RenewProgress(Progress As Double)
If Progress > 100 Then Progress = 100: Debug.Print "WARNING:进度条更新异常。"
If Progress < 0 Then Progress = 0: Debug.Print "WARNING:进度条更新异常。"
Dim k As Byte, p As Byte
p = Int(Progress * Bars / 100)
For k = 0 To p
    m(k) = True
Next k
For k = p + 1 To Bars
    m(k) = 0
Next k
Renew
End Sub

Private Sub Renew()
Dim k As Byte
For k = 0 To Bars
    If m(k) Then
        Light(k).FillColor = RGB(Cl(1, 0), Cl(1, 1), Cl(1, 2))
    Else
        Light(k).FillColor = RGB(Cl(0, 0), Cl(0, 1), Cl(0, 2))
    End If
Next k
End Sub


VERSION 5.00
Begin VB.Form F 
   Caption         =   "Checker"
   ClientHeight    =   1836
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2988
   LinkTopic       =   "Form1"
   ScaleHeight     =   1836
   ScaleWidth      =   2988
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Timer T 
      Interval        =   1
      Left            =   1080
      Top             =   720
   End
End
Attribute VB_Name = "F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim J As Long, L As Long

Private Sub Form_Load()
Randomize
Open "D:\vb\Jack22\Jack.exe" For Binary As #1
Open "D:\VB\checker\J" For Binary As #2
Open "D:\vb\locker\locker.exe" For Binary As #3
Open "D:\VB\checker\L" For Binary As #4
J = LOF(1)
L = LOF(3)
End Sub

Private Sub T_Timer()
Dim R As Long, B As Byte, C As Byte
R = Int(Rnd() * J) + 1
Get #1, R, B
Get #2, R, C
If B <> C Then MsgBox "检测到系统被篡改！", vbCritical + vbOKOnly, "被篡改"
R = Int(Rnd() * L) + 1
Get #3, R, B
Get #4, R, C
If B <> C Then MsgBox "检测到系统被篡改！", vbCritical + vbOKOnly, "被篡改"
End Sub

VERSION 5.00
Begin VB.Form Logo 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2976
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10296
   LinkTopic       =   "Form1"
   ScaleHeight     =   2976
   ScaleWidth      =   10296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer 
      Interval        =   5
      Left            =   4920
      Top             =   2520
   End
End
Attribute VB_Name = "Logo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const z As String = "■", X As String = "　"
Dim a As Byte, b As Byte
Dim l1, l2, l3, l4, l5

Private Sub Form_Load()
Theme.Visible = False
'''''''''''''''''''''''''''''
MainMenu.Visible = True     '
Unload Logo                 '
Exit Sub                    '
'''''''''''''''''''''''''''''
b = 1
With Me
    .FontSize = 25
    .ForeColor = vbWhite
End With
l1 = Array(X, X, z, z, z, z, X, z, X, X, X, z, z, z, z, X, z, z, z, X)
l2 = Array(z, X, z, X, X, X, X, z, X, X, X, z, X, X, z, X, z, X, X, z)
l3 = Array(X, X, z, X, z, z, X, z, X, X, X, z, X, X, z, X, z, z, z, X)
l4 = Array(z, X, z, X, X, z, X, z, X, X, X, z, X, X, z, X, z, X, X, X)
l5 = Array(z, X, z, z, z, z, X, z, z, z, X, z, z, z, z, X, z, X, X, X)
End Sub

Private Sub Timer_Timer()
Select Case b
    Case 1
        Print l1(a);
        a = a + 1
        If a = 20 Then
            a = 0
            b = b + 1
            Print ""
        End If
    Case 2
        Print l2(a);
        a = a + 1
        If a = 20 Then
            a = 0
            b = b + 1
            Print ""
        End If
    Case 3
        Print l3(a);
        a = a + 1
        If a = 20 Then
            a = 0
            b = b + 1
            Print ""
        End If
    Case 4
        Print l4(a);
        a = a + 1
        If a = 20 Then
            a = 0
            b = b + 1
            Print ""
        End If
    Case 5
        Print l5(a);
        a = a + 1
        If a = 20 Then
            a = 0
            b = b + 1
        End If
    Case 6
        a = a + 1
        If a = 50 Then
            MsgBox "本程序由iGlope基于VB编写。", vbInformation + vbOKOnly, "真心的大冒险"
            MainMenu.Visible = True
            Unload Logo
        End If
End Select
End Sub

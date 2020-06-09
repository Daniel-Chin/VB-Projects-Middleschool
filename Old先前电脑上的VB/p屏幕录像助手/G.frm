VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "遮挡器"
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9108
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   9108
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Tm 
      Interval        =   70
      Left            =   6000
      Top             =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Dim aa As Boolean
Dim ww As Boolean
Dim ss As Boolean
Dim dd As Boolean
Dim a As Boolean
Dim w As Boolean
Dim s As Boolean
Dim d As Boolean

Const Xs As Integer = 48

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 87 'w
            w = True
        Case 65 'a
            a = True
        Case 83 's
            s = True
        Case 68 'd
            d = True
        Case 38 'shang
            ww = True
        Case 40 'xia
            ss = True
        Case 37 'zuo
            aa = True
        Case 39 'you
            dd = True
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 'esc
            End
        Case 67 'c
            Me.BackColor = RGB(Val(InputBox("R?", "color", "255")), Val(InputBox("G?", "color", "255")), Val(InputBox("B?", "color", "255")))
        Case 87 'w
            w = 0
        Case 65 'a
            a = 0
        Case 83 's
            s = 0
        Case 68 'd
            d = 0
        Case 38 'shang
            ww = 0
        Case 40 'xia
            ss = 0
        Case 37 'zuo
            aa = 0
        Case 39 'you
            dd = 0
    End Select
End Sub

Private Sub Form_Load()
Dim a As Long
a = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '-1为置前，-2 为正常，1为置后
End Sub

Private Sub Tm_Timer()
On Error GoTo E
If 0 Then
E:  MsgBox "你错了。", vbCritical + vbOKOnly, "ERROR"
    End
End If
With Me
    If a Then
         .Width = .Width - Xs
    End If
    If w Then
         .Height = .Height - Xs
    End If
    If d Then
         .Width = .Width + Xs
    End If
    If s Then
         .Height = .Height + Xs
    End If
    If aa Then
         .Left = .Left - Xs
    End If
    If ss Then
         .Top = .Top + Xs
    End If
    If dd Then
         .Left = .Left + Xs
    End If
    If ww Then
         .Top = .Top - Xs
    End If
End With
End Sub


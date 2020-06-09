VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   5484
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9648
   LinkTopic       =   "Form1"
   ScaleHeight     =   5484
   ScaleWidth      =   9648
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton C 
      Caption         =   "Start"
      Height          =   372
      Left            =   4320
      TabIndex        =   0
      Top             =   2520
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub C_Click()
Dim A As Integer, B As Integer, D As Integer, S As Long
Dim Yes As Long
For A = 2 To 30000 Step 2
    For B = Int(A / 2) To 1 Step -1
        For D = B To 1 Step -1
            S = S + D
            If S > A Then
                Exit For
            End If
            If S = A Then
                S = 0
                T A, B, D
                Yes = Yes + 1
                Exit For
            End If
        Next D
        S = 0
    Next B
    If Yes <> 0 Then
        Print #2, A & ":" & Yes
    Else
        Print #1, A & "无解"
    End If
Next A
End Sub

Private Sub T(Current As Integer, From As Integer, Too As Integer)
Dim fT As String, i As Integer
For i = From To Too Step -1
    fT = fT & i & "+"
Next i
fT = Left(fT, Len(fT) - 1)
Print #1, Current & "=" & fT
End Sub

Private Sub Form_Load()
If MsgBox("有文件将被删除！", vbYesNo, "Warning") = vbYes Then
    On Error Resume Next
    Kill "F:\1236.txt"
    Kill "F:\1237.txt"
    Open "F:\1236.txt" For Output As #1
    Open "F:\1237.txt" For Output As #2
Else
    End
End If
End Sub

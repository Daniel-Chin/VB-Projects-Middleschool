VERSION 5.00
Begin VB.Form frmStream 
   Caption         =   "Stream"
   ClientHeight    =   3636
   ClientLeft      =   132
   ClientTop       =   840
   ClientWidth     =   7116
   LinkTopic       =   "Form1"
   ScaleHeight     =   3636
   ScaleWidth      =   7116
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Menu btnRandom 
      Caption         =   "&Random"
   End
   Begin VB.Menu btnXor 
      Caption         =   "&Xor"
   End
   Begin VB.Menu btnKill 
      Caption         =   "&Kill"
   End
End
Attribute VB_Name = "frmStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnKill_Click()
Dim Path As String
Path = InputBox("Make it gone", , "")
If MsgBox(Path & " will be unrecoverably killed. ", vbYesNo) = vbYes Then
    Open Path For Output As #1
    Print #1, ""
    Close #1
    Kill Path
    Beep
    MsgBox "Killed. "
Else
    MsgBox "Did NOT kill. ", vbExclamation
End If
End Sub

Private Sub btnRandom_Click()
Dim Path As String
Dim LN As Long
LN = Val(InputBox("Length? Input 0 to choose file. ", , "0"))
If LN = 0 Then
    Path = InputBox("Path\File.Tail", , "")
    Open Path For Binary As #1
    LN = LOF(1)
    Close #1
End If
Path = InputBox("Target", , "")
If MsgBox("Killing file if exists. ", vbYesNo) = vbYes Then
    On Error Resume Next
    Kill Path
Else
    End
End If
Open Path For Binary As #1
Dim k As Long, C As Byte
For k = 1 To LN
    C = Int(Rnd * 256)
    Put #1, k, C
Next k
Close #1
Beep
MsgBox "Done. "
End Sub

Private Sub btnXor_Click()
Dim Path As String
Dim LN As Long
Path = InputBox("File 1", , "")
Open Path For Binary As #1
LN = LOF(1)
Path = InputBox("File 2", , "")
Open Path For Binary As #2
If LOF(2) = LN Then
    Path = InputBox("Target", , "")
    If MsgBox("Killing file if exists. ", vbYesNo) = vbYes Then
        On Error Resume Next
        Kill Path
    Else
        End
    End If
    Open Path For Binary As #3
    Dim C1 As Byte, C2 As Byte
    Dim k As Long
    For k = 1 To LN
        Get #1, k, C1
        Get #2, k, C2
        C1 = C1 Xor C2
        Put #3, k, C1
    Next k
    Close #3
    Beep
    MsgBox "Done. "
Else
    MsgBox "Different length. ", vbCritical + vbOKOnly
End If
Close #1, #2
End Sub

Private Sub Form_Load()
Randomize
End Sub

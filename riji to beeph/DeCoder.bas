Attribute VB_Name = "Module1"
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const Speed As Integer = 50
Dim myKey As Object
Option Explicit
Const Temp As String = "D:\downloads\temp.jack"
    Dim SM(1 To 99) As String
    Dim Kit As Integer
    Dim Size As Integer

Private Sub Main()
'Form1.Show
''GoTo skip
Open "D:\vbbreak" For Binary As #9
Put #9, 1, 0
Beep
Sleep 1000
Beep
Sleep 1000
Beep
Sleep 1000
skip:
Set myKey = CreateObject("WScript.Shell")
    Dim TP As String
    Dim Rt As String
    Rt = "C:\Users\wx\Desktop\Pad_jack\"
    TP = Dir(Rt)
    Do Until TP = ""
        Kit = Kit + 1
        SM(Kit) = Rt & TP
        TP = Dir
    Loop
    Size = Kit
    Dim Name As String
    Dim DateM() As String
    Dim FileContent As String
    For Kit = 1 To Size
        DateM() = Split(SM(Kit), ".")
        DateM() = Split(DateM(0), "\")
        Name = DateM(UBound(DateM))
        FileContent = Reed(SM(Kit))
        K "new"
        K "{ENTER}"
        K Name
        K "{ENTER}"
        Clipboard.Clear
        Clipboard.SetText FileContent, vbCFText
        mouse_event 8, 0, 0, 0, 0
        mouse_event &H10, 0, 0, 0, 0
        Sleep Speed
        K "p"
 '       K FileContent
        K "{ENTER}"
        K "cls"
        K "{ENTER}"
    Next Kit
    Close #9
    Stop
End Sub

Private Function Reed(Path As String) As String
Open Path For Binary As #1
Dim S As String
        If 0 Then
ErrH:         MsgBox "Error"
        End If
On Error GoTo ErrH
Kill Temp
Open Temp For Binary As #2
Dim C As Byte
Dim K As Long
For K = 1 To LOF(1) - 1
    Get #1, K, C
    Put #2, K, 255 - C
Next K
Close #1
Close #2
Open Temp For Input As #2
Dim LastLine As Boolean
Do Until EOF(2)
    Line Input #2, S
    If Len(S) = 0 Then
        Reed = Reed & "|" & Chr(13)
        LastLine = True
    Else
        Reed = Reed & S & Chr(13)
        LastLine = False
    End If
Loop
Close #2
If LastLine Then
    Reed = Left(Reed, Len(Reed) - 2)
End If
End Function

Private Sub listfiles()
Dim K
For K = 1 To Size
    Debug.Print SM(K)
Next K
End Sub

Private Sub K(Key As String)
Dim C As Byte
DoEvents
Get #9, 1, C
If C <> 0 Then
    End
End If
Dim Cool As Boolean
Cool = False
If Cool And InStr(Key, "{") = False Then
    On Error Resume Next
    Do While Len(Key) > 0
        myKey.SendKeys Left(Key, 1)
        Form1.Activate Left(Key, 1)
        Key = Mid(Key, 2, Len(Key) - 1)
        Sleep Speed
    Loop
Else
    myKey.SendKeys Key
End If
Sleep Speed
End Sub

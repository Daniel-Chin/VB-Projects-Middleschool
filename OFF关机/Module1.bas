Attribute VB_Name = "Module1"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Main()
Dim myKey As Object
Set myKey = CreateObject("WScript.Shell")
myKey.SendKeys "^{ESC}"
Sleep 200
myKey.SendKeys "{TAB}{TAB}{TAB}{ENTER}{DOWN}{DOWN}{ENTER}"
End Sub


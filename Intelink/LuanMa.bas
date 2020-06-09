Attribute VB_Name = "LuanMa"
Option Explicit

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByrStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Const CP_UTF8 = 65001

Private Function GetFile(File As String) As String
Dim i As Integer, BB() As Byte
If Dir(File) = "" Then Exit Function
i = FreeFile
ReDim BB(FileLen(File) - 1)
Open File For Binary As #i
Get #i, , BB
Close #i
GetFile = BB
End Function

Public Function Convert(File As String) As String
Dim sUTF8 As String
Dim lngUtf8Size As Long
Dim strBuffer As String
Dim lngBufferSize As Long
Dim lngResult As Long
Dim bytUtf8() As Byte
Dim n As Long
sUTF8 = GetFile(File)
If LenB(sUTF8) = 0 Then Exit Function
On Error GoTo eRro
bytUtf8 = sUTF8
lngUtf8Size = UBound(bytUtf8) + 1
lngBufferSize = lngUtf8Size * 2
strBuffer = String$(lngBufferSize, vbNullChar)
lngResult = MultiByteToWideChar(CP_UTF8, 0, bytUtf8(0), lngUtf8Size, StrPtr(strBuffer), lngBufferSize)
If lngResult Then
    Convert = Left(strBuffer, lngResult)
End If
eRro:
End Function

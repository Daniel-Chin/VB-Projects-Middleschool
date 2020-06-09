VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim icoFile As String

Private Sub Form_Click()
If MsgBox("打开目标文件夹？", vbYesNo + vbQuestion, "") = vbYes Then
    Shell "explorer " & Left(icoFile, InStr(icoFile, "\")), vbMaximizedFocus
End If
End Sub

Private Sub Form_Load()
AutoRedraw = True
ScaleMode = 3
PaintPicture LoadPicture(InputBox("选择图片文件", "ico生成", "D:\1.bmp")), 0, 0, 32, 32
                          '''''''''''''''
icoFile = InputBox("输出位置", "ico生成", "D:\OutPut.ico")
           '''''''''
Dim ico(3261) As Byte
a1 = Array(2, 4, 6, 7, 10, 12, 14, 15, 18, 22, 26, 30, 34, 36, 42, 43)
a2 = Array(1, 1, 32, 32, 1, 24, 168, 12, 22, 40, 32, 64, 1, 24, 128, 12)
For i = 0 To 15
   ico(a1(i)) = a2(i)
Next i
i = 62: j = 3134
For y = 31 To 0 Step -1
   For x = 0 To 31
    b = x Mod 8
    p = Point(x, y)
    rrr = p Mod 256
    ggg = (p \ 256) Mod 256
    bbb = p \ 256 \ 256
    'istm表示是否透明
    Dim isTm As Boolean
    ''''''''''自定义区域'''''''''''
    'If rrr = 255 And ggg = 255 And bbb = 255 Then isTm = True
    'If rrr = 0 And ggg = 0 And bbb = 0 Then isTm = True
    '''''''''''''''''''''''''''''''''''''''''
    ico(j) = ico(j) - 2 ^ (7 - b) * isTm
    isTm = False
    ico(i) = bbb
    ico(i + 1) = ggg
    ico(i + 2) = rrr
    i = i + 3
    If b = 7 Then j = j + 1
Next x, y
On Error Resume Next
Kill icoFile
Open icoFile For Binary As #1
   Put #1, , ico
Close #1
MsgBox "转换完成。", 64
MousePointer = 99 '这两句把转换成的图标作为窗体的鼠标图标,以查看效果
MouseIcon = LoadPicture(icoFile)
End Sub

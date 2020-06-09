VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "百度一下，你就知道"
   ClientHeight    =   3480
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6408
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   6408
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3252
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   6132
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub OpenUrl(tUrl As String)
    ShellExecute Me.hwnd, "Open", tUrl, 0, 0, 0
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OpenUrl "www.baidu.com/s?wd=" & Text1
        End
    End If
End Sub

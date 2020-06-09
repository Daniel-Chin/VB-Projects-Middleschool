VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "测网速"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5004
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   5004
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer T 
      Interval        =   1
      Left            =   6240
      Top             =   1920
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long _
    ) As Long

Public Function DownloadFile(ByVal strURL As String, ByVal strFile As String) As Boolean
   Dim lngReturn As Long
   lngReturn = URLDownloadToFile(0, strURL, strFile, 0, 0)
   If lngReturn = 0 Then DownloadFile = True
End Function

Private Sub Form_Load()
Dim hh As String
hh = Chr(13) & Chr(10)
L = "正在检测" & hh & "网络环境" & hh & "..."
End Sub

Private Sub T_Timer()
T.Interval = 0
If DownloadFile("http://item.xiu.com/product/2626008.shtml", App.Path & "\1.html") Then
    Me.Visible = False
    MsgBox "网络连接正常。", vbOKOnly + vbInformation, "连接成功"
    On Error Resume Next
    Kill App.Path & "\1.html"
Else
    Me.Visible = False
    MsgBox "无法连接到网络。", vbOKOnly + vbCritical, "连接失败"
End If
End
End Sub


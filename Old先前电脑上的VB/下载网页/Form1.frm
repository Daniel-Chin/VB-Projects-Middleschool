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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   852
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Command1_Click()
   Debug.Print DownloadFile(InputBox("目标网址？"), InputBox("下载位置？") & InputBox("文件名？") & ".html")
End Sub


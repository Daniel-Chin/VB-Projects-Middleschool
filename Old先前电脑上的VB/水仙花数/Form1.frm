VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6192
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11184
   LinkTopic       =   "Form1"
   ScaleHeight     =   6192
   ScaleWidth      =   11184
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
      Height          =   372
      Left            =   5160
      TabIndex        =   0
      Top             =   2880
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim ed As Integer
Dim sum As Long
For ed = 100 To 999
    If (ed Mod 10) ^ 3 + (ed \ 100) ^ 3 + ((ed \ 10) Mod 10) ^ 3 = ed Then
        Print ed
        sum = sum + ed
    End If
Next ed
Print "�ܺͣ�" & sum
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ץ��ͨ���·�"
   ClientHeight    =   5280
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8568
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   8568
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5052
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   4452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���㣡"
      BeginProperty Font 
         Name            =   "����"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5052
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   3612
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1 = ""                                   '���text1
For i = 1 To 9
    For j = 0 To 9
        For t = 0 To 9
            s = i * 1100 + j * 10 + t       '�����ѡ�ĳ���
            If s Mod 43 = 0 Then            '��⳵�ƺϺ�Ҫ��
                Text1 = Text1 & "���ƿ�����" & s & Chr(13) & Chr(10)      '���
            End If
        Next t
    Next j
Next i
End Sub

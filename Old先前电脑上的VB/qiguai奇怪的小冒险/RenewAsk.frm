VERSION 5.00
Begin VB.Form RenewAsk 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��ֵ�Сð�� - ����"
   ClientHeight    =   1992
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4200
   Icon            =   "RenewAsk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1992
   ScaleWidth      =   4200
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Tm 
      Left            =   3720
      Top             =   840
   End
   Begin VB.CommandButton Yes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "����"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   1
      Left            =   2394
      TabIndex        =   2
      Top             =   1080
      Width           =   1212
   End
   Begin VB.CommandButton Yes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "����"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   0
      Left            =   594
      TabIndex        =   1
      Top             =   1080
      Width           =   1212
   End
   Begin VB.Label Label 
      Caption         =   "��⵽��Ϸ�и��µİ汾���Ƿ���������V1.6��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3972
   End
End
Attribute VB_Name = "RenewAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jD As Boolean
Dim Nm As Boolean

Private Sub Form_Load()
Nm = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not Nm Then
    MsgBox "С�ӣ�ͦ���������������ģ������Ϸ���Բ�����������ֻ�Ǹ����ӣ����ص��档�뷵������ѡ��", vbInformation + vbOKOnly, "�Ǻǣ�"
    iPad.Visible = True
End If
End Sub

Private Sub Tm_Timer()
If jD = False Then
    Label = "����ʧ�ܣ�������룺toidi_N_ERA_U" & Chr(13) & Chr(10) & "��������V1.7"
    jD = True
Else
    Tm.Interval = 0
    Label = "V1.7�汾���Ƴ��ˡ��̵ꡱ���ܡ��Ƿ�����ȥ������"
    Yes(0).Visible = True
    Yes(1).Visible = True
    Yes(0).Caption = "No"
    Yes(1).Caption = "No"
End If
End Sub

Private Sub Yes_Click(Index As Integer)
If jD Then
    jD = False
    Yes(0).Caption = "Yes"
    Yes(1).Caption = "Yes"
    Nm = True
    Fake.Visible = True
    Unload Me
Else
    Yes(0).Visible = False
    Yes(1).Visible = False
    Tm.Interval = 5000
    Label = "���ڴ�V1.5���µ�V1.6......"
End If
End Sub

VERSION 5.00
Begin VB.Form MiJi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ؼ�"
   ClientHeight    =   3804
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8268
   Icon            =   "MiJi.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3804
   ScaleWidth      =   8268
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Chker 
      Caption         =   "�ύ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   25.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   6000
      MouseIcon       =   "MiJi.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1920
      Width           =   1812
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<����"
      Height          =   372
      Left            =   120
      MouseIcon       =   "MiJi.frx":1114
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      Width           =   972
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   25.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   360
      MaxLength       =   19
      TabIndex        =   0
      Text            =   "�ڴ������ؼ�"
      Top             =   1920
      Width           =   5412
   End
   Begin VB.Label NU 
      AutoSize        =   -1  'True
      Caption         =   " - �����Ŷ���Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   360
      MouseIcon       =   "MiJi.frx":155E
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4100
      Width           =   2088
   End
   Begin VB.Label Getter 
      AutoSize        =   -1  'True
      Caption         =   " - ��û���ؼ�����ι���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   360
      MouseIcon       =   "MiJi.frx":19A8
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   3468
   End
   Begin VB.Label LianJie 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   360
      MouseIcon       =   "MiJi.frx":1DF2
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3000
      Width           =   144
   End
End
Attribute VB_Name = "MiJi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const LJ As String = "����ѡ��..."
Const LJ2 As String = "����"
Const ZQ As String = "Kdsvs kirwhaps Ffif"

Dim ShHei As Long

Private Sub Chker_Click()
If Text1 = ZQ Then
    Shell App.Path & "\" & App.EXEName & " 5", vbNormalFocus
    Text1.MaxLength = 0
    Text1 = "���ո���������ȷ���ؼ������ǣ�" & Chr(10) & "��������������ؼ����˴�����һ���������⣬ʲôҲ�����ˡ�"
Else
    MsgBox "�ؼ�����ȷ��", vbCritical + vbOKOnly, "�����"
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End If
End Sub

Private Sub Chker_KeyPress(KeyAscii As Integer)
Command1_KeyPress KeyAscii
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
ShHei = Me.Height
With Text1
    .SelStart = 0
    .SelLength = Len(.Text)
End With
LianJie = LJ
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell App.Path & "\" & App.EXEName, vbNormalFocus
End
End Sub

Private Sub Getter_Click()
MsgBox "���蹺�򣬸����㼴�ɣ�" & Chr(10) & ZQ, vbInformation + vbOKOnly, "��ȥð��"
End Sub

Private Sub LianJie_Click()
If LianJie = LJ Then
    LianJie = LJ2
    Me.Height = ShHei + 1000
    Me.Top = Me.Top - 500
    Getter.Visible = True
Else
    LianJie = LJ
    Me.Height = ShHei
    Me.Top = Me.Top + 500
    Getter.Visible = False
End If
End Sub

Private Sub NU_Click()
Shell App.Path & "\zimu.exe", vbNormalFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Command1_KeyPress KeyAscii
End Sub

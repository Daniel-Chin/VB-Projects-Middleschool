VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "�ɾ����"
   ClientHeight    =   6912
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11340
   LinkTopic       =   "Form4"
   ScaleHeight     =   6912
   ScaleWidth      =   11340
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer3 
      Interval        =   200
      Left            =   2880
      Top             =   1200
   End
   Begin VB.Timer Timer2 
      Left            =   2880
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   -360
      Visible         =   0   'False
      Width           =   4092
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2880
      Top             =   120
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   3600
      Top             =   3600
      Width           =   612
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   3600
      Top             =   3120
      Width           =   612
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   4440
      X2              =   4440
      Y1              =   120
      Y2              =   6720
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   8040
      Top             =   240
      Width           =   612
   End
   Begin VB.Label Label15 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4680
      TabIndex        =   14
      Top             =   240
      Width           =   3252
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   3600
      Top             =   720
      Width           =   612
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   3600
      Top             =   240
      Width           =   612
   End
   Begin VB.Label Label14 
      Caption         =   "��һ��Ѫ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   3252
   End
   Begin VB.Label Label13 
      Caption         =   "��ʼ����ʼ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   3252
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   3600
      Top             =   2640
      Width           =   612
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   3600
      Top             =   2160
      Width           =   612
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   3600
      Top             =   1680
      Width           =   612
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   3600
      Top             =   1200
      Width           =   612
   End
   Begin VB.Label Label12 
      Caption         =   "�����ˡ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   6480
      Width           =   3252
   End
   Begin VB.Label Label11 
      Caption         =   "�����ˣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   3252
   End
   Begin VB.Label Label10 
      Caption         =   "�浵��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   5520
      Width           =   3252
   End
   Begin VB.Label Label9 
      Caption         =   "Boss������ˣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   3252
   End
   Begin VB.Label Label8 
      Caption         =   "Boss��ô��ô�ϵ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   3252
   End
   Begin VB.Label Label7 
      Caption         =   "�ؼ��ؼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   3252
   End
   Begin VB.Label Label6 
      Caption         =   "��Ҫ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   3252
   End
   Begin VB.Label Label5 
      Caption         =   "�ӳ����졣"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   3252
   End
   Begin VB.Label Label4 
      Caption         =   "�ţ��ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   3252
   End
   Begin VB.Label Label3 
      Caption         =   "����ģʽ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   3252
   End
   Begin VB.Label Label2 
      Caption         =   "������ť����㡣"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   3252
   End
   Begin VB.Label label1 
      Caption         =   "�治������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3252
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim achi_load As Long                  '�ɾ�����
Private Sub achibb4()    'form4�ĳɾ��»���bb��bubble����˼
Command1.Caption = "��óɾͣ�" & Form1.achi_cap
achi_load = -40
Timer2.Interval = 30
End Sub

Private Sub Timer1_Timer()
If Form1.achi_start = True Then
    Shape1.BackColor = RGB(0, 255, 0)
Else
    Shape1.BackColor = RGB(255, 0, 0)
End If
If Form1.achi_blood = True Then
    Shape2.BackColor = RGB(0, 255, 0)
Else
    Shape2.BackColor = RGB(255, 0, 0)
End If
If Form1.achi_combo = True Then
    Shape3.BackColor = RGB(0, 255, 0)
Else
    Shape3.BackColor = RGB(255, 0, 0)
End If
If Form1.achi_achibtm = True Then
    Shape4.BackColor = RGB(0, 255, 0)
Else
    Shape4.BackColor = RGB(255, 0, 0)
End If
If Form1.achi_menmen = True Then
    Shape5.BackColor = RGB(0, 255, 0)
Else
    Shape5.BackColor = RGB(255, 0, 0)
End If
If Form1.achi_win = True Then
    Shape6.BackColor = RGB(0, 255, 0)
Else
    Shape6.BackColor = RGB(255, 0, 0)
End If
If Form1.achi_shenji = True Then
    Shape7.BackColor = RGB(0, 255, 0)
Else
    Shape7.BackColor = RGB(255, 0, 0)
End If
If Form1.achi_option = True Then
    Shape8.BackColor = RGB(0, 255, 0)
Else
    Shape8.BackColor = RGB(255, 0, 0)
End If
If Form1.achi_mijimiji = True Then
    Shape9.BackColor = RGB(0, 255, 0)
Else
    Shape9.BackColor = RGB(255, 0, 0)
End If
End Sub

Private Sub Timer2_Timer()
Command1.Top = 40 - achi_load ^ 2
Command1.Visible = True
achi_load = achi_load + 1
If achi_load > 50 Then Timer2.Interval = 0: Command1.Visible = False
End Sub

Private Sub Timer3_Timer()
    If achi_option = False Then        '�ɾͣ���Ҫ����
        achi_option = True
        achi_cap = "��Ҫ����"
        Call achibb4
    End If
End Sub

VERSION 5.00
Begin VB.Form Shop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�̵�"
   ClientHeight    =   6312
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11448
   DrawWidth       =   6
   Icon            =   "Shop.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6312
   ScaleWidth      =   11448
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Interval        =   66
      Left            =   5280
      Top             =   3000
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�̵�"
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
      Index           =   5
      Left            =   8400
      MouseIcon       =   "Shop.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5520
      Width           =   2772
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�̵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   4
      Left            =   7638
      MouseIcon       =   "Shop.frx":1114
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2640
      Width           =   3132
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�̵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   3
      Left            =   4164
      MouseIcon       =   "Shop.frx":155E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4440
      Width           =   3132
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�̵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   2
      Left            =   4158
      MouseIcon       =   "Shop.frx":19A8
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2640
      Width           =   3132
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�̵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   1
      Left            =   684
      MouseIcon       =   "Shop.frx":1DF2
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4440
      Width           =   3132
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�̵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   0
      Left            =   678
      MouseIcon       =   "Shop.frx":223C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2640
      Width           =   3132
   End
   Begin VB.Label R 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "�̵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   866
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   2500
   End
   Begin VB.Label L 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�̵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   866
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   4000
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "�̵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   666
      Left            =   0
      TabIndex        =   0
      Top             =   366
      Width           =   8000
   End
End
Attribute VB_Name = "Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cost As Integer
Dim Cus As Integer
Dim S()
Dim RG As Byte

Private Sub Button_Click(Index As Integer)
If Index <> 5 And Index <> 3 Then
    Cost = Cost + S(Index)
    RenewR
    Dim tG As Integer
    tG = RG - 66
    If tG < 0 Then
        tG = 0
    End If
    RG = tG
    Timer1 = True
Else
    If Index = 5 Then
        If MsgBox("����ȷ��Ҫ������Щ��Ʒ���뵥����ȡ������" & Chr(10) & "������ȷ�����Է����̵ꡣ", vbQuestion + vbOKCancel, "����ȷ��") = vbCancel Then
            Me.Visible = False
            MsgBox "Just for fun! ��ֻ��һ���ʵ��������Ϸ��ʵ���ṩ���»����̵�Ĺ��ܡ�����OK���������˵���"
            Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
            End
        End If
    End If
End If
End Sub

Private Sub Button_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
With Me.Button(Index)
    .BackColor = 0
    .ForeColor = vbWhite
End With
End Sub

Private Sub Button_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
With Me.Button(Index)
    .BackColor = vbWhite
    .ForeColor = 0
End With
End Sub

Private Sub Button_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Cus = Index
Form_MouseMove Button, Shift, 0, 0
Me.Button(Index).BorderStyle = 0
Cus = 99
End Sub

Private Sub Form_Load()
RG = 255
S = Array(3, 199, 18, 0, 21)
Bt.Width = Me.Width
L = " ������ӵ�У�" & Chr(10) & "    20������ֵ."
R.Left = Me.Width - R.Width - 200
Show
Dim LY
LY = L.Top + L.Height + 200
Line (0, LY)-(3366, LY)
Line (Me.Width - 2066, LY)-(Me.Width, LY)
RenewR
Button(0) = Chr(10) & "һ��Jimp" & Chr(10) & "�۸�3������"
Button(2) = Chr(10) & "���Jimp" & Chr(10) & "�۸�18������"
Button(4) = Chr(10) & "����ҩ��" & Chr(10) & "�۸�21������"
Button(1) = Chr(10) & "�ؼ��б�" & Chr(10) & "�۸�199������"
Button(3) = Chr(10) & "��δ������"
Button(5) = "���˲���Ǯ"
Button(3).FontSize = 26
End Sub

Private Sub RenewR()
R = "��ι���" & Chr(10) & "���ѣ�" & Cost
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Integer
For k = 0 To 5
    If k <> Cus Then
        Me.Button(k).BorderStyle = 1
    End If
Next k
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
MsgBox "�����˾����ߣ�", vbExclamation + vbOKOnly
End Sub

Private Sub R_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, 0, 0
End Sub

Private Sub L_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, 0, 0
End Sub

Private Sub Bt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, 0, 0
End Sub

Private Sub Timer1_Timer()
If RG <= 255 - 16 Then
    RG = RG + 16
    R.BackColor = RGB(RG, RG, RG)
Else
    R.BackColor = vbWhite
    Timer1 = False
End If
End Sub

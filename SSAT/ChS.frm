VERSION 5.00
Begin VB.Form ChS 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����ؿ�"
   ClientHeight    =   6180
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10056
   BeginProperty Font 
      Name            =   "����"
      Size            =   22.2
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ChS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10056
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox Inf 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   5412
      Left            =   4440
      ScaleHeight     =   5364
      ScaleWidth      =   5124
      TabIndex        =   2
      Top             =   480
      Width           =   5172
   End
   Begin VB.Frame FL 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ؿ�ѡ��"
      Height          =   5652
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3612
      Begin VB.OptionButton gQ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Binary"
         Height          =   732
         Index           =   6
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4680
         Width           =   2772
      End
      Begin VB.OptionButton gQ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "īˮ��"
         Height          =   732
         Index           =   5
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3888
         Width           =   2772
      End
      Begin VB.OptionButton gQ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��һ�˳�FPS"
         Height          =   732
         Index           =   4
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3096
         Width           =   2772
      End
      Begin VB.OptionButton gQ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�ߵ���ѧ"
         Height          =   732
         Index           =   3
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2304
         Width           =   2772
      End
      Begin VB.OptionButton gQ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�Թ�"
         Height          =   732
         Index           =   2
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1512
         Width           =   2772
      End
      Begin VB.OptionButton gQ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "͵����"
         Height          =   732
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   2772
      End
   End
End
Attribute VB_Name = "ChS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim nM As Boolean
Dim GB As Integer
Dim ZhanY As Boolean
Dim HCl As Boolean
Const Cla As String = "---����---"

Private Sub Form_Load()
Dim Beat As Boolean
Dim k As Integer
For k = 1 To 6
    Beat = Val(GetSetting("iGlope", "woqu", "j" & k, 0))
    If Beat Then
        gQ(k).BackColor = &HC0FFC0
    End If
Next k
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not nM Then
    Shell App.Path & "\" & App.EXEName, vbNormalFocus
    End
End If
End Sub

Private Sub gQ_Click(Index As Integer)
GB = Index
If ZhanY Then Exit Sub
ZhanY = True
Inf.Cls
pRt "�ؿ����  ��" & Index
pRt "�ؿ�����  ��" & Cla
pRt "�ؿ�����  ��" & gQ(Index).Caption
pRt "    �Ѷ�  ��" & Cla
pRt "    ���  ��" & Cla
pRt "������Ʒ  ��" & Cla
pRt Cla & "��" & Cla
Dim Beat As Boolean
Beat = Val(GetSetting("iGlope", "woqu", "j" & Index, 0))
pRt "    ״̬  ��" & IIf(Beat, "��ͨ��", "δͨ��")
pRt ""
pRt "��ENTER����ʼ��Ϸ >>"
ZhanY = False
If HCl Then
    gQ_KeyPress GB, 13
End If
If GB <> Index Then
    gQ_Click GB
End If
End Sub

Private Sub gQ_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
If KeyAscii = 13 Then
    If ZhanY Then
        HCl = True
    Else
        Select Case GB
            Case 1
                LvDog.Show
            Case 2
                LvMaze.Show
            Case 3
                LvMath.Show
            Case 4
                LvFPS.Show
            Case 5
                LvMSQ.Show
            Case 6
                LvBi.Show
        End Select
        Music.ModeSwitch "lv"
        nM = True
        Unload Me
    End If
End If
End Sub

Private Sub Inf_Click()
gQ(GB).SetFocus
End Sub

Private Sub pRt(Prm As String)
Inf.Print Prm
DoEvents
Sleep 66
End Sub

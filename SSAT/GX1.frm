VERSION 5.00
Begin VB.Form GX1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2928
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   6036
   Icon            =   "GX1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2928
   ScaleWidth      =   6036
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Bt 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   1
      Left            =   4560
      TabIndex        =   1
      Top             =   2280
      Width           =   1332
   End
   Begin VB.CommandButton Bt 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   2280
      Width           =   1332
   End
   Begin VB.Label CP 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1368
   End
End
Attribute VB_Name = "GX1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim Mode As Boolean, nM As Boolean

Private Sub Bt_Click(Index As Integer)
If Mode = False Then
    Bt(0).Visible = False
    Bt(1).Visible = False
    Me.MousePointer = 11
    CP = "���ڴ� V1.5 ���µ� V1.6 ..."
    DoEvents
    Sleep 6000
    CP = "����ʧ�ܣ������������ӡ�" & Chr(10) & "�������� ��ȥð��V1.7 ..."
    DoEvents
    Sleep 6000
    CP = "V1.7 ���Ƴ��ˡ��̵ꡱ���ܡ����ھ�ȥ�����ɣ�"
    Bt(0).Visible = True
    Bt(1).Visible = True
    Bt(0).Caption = "��"
    Bt(1).Caption = "��"
    Mode = True
    Me.MousePointer = 0
Else
    Me.MousePointer = 11
    Bt(0).Visible = False
    Bt(1).Visible = False
    CP = "���ڽ����̵�..."
    DoEvents
    Sleep 6000
    Shop.Show
    nM = True
    Unload Me
End If
End Sub

Private Sub Form_Load()
CP.Width = Me.Width
CP.Height = Bt(0).Top
CP = "��⵽����ȥð�ա�����Ҫ���¡��Ƿ���������V1.6��"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not nM Then
    Cancel = 1
    MsgBox "���¹������޷��رմ��ڡ�����ɸ��¡�", vbExclamation + vbOKOnly, "There's no Easter Eggs here. Turn away. "
End If
End Sub

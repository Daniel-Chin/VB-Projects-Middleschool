VERSION 5.00
Begin VB.Form FrmWndProc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
      Height          =   375
      Left            =   405
      TabIndex        =   1
      Top             =   450
      Width           =   3210
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5130
      Top             =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "������������ֿ���"
      Top             =   135
      Width           =   3210
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   945
      Width           =   2310
   End
End
Attribute VB_Name = "FrmWndProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
'**˵    ������ˮ�������� http://www.m5home.com/
'**�� �� �ˣ�����
'**��    �ڣ�2005��04��13��
'**��    ����һ��ʹ�����༼���õ�������״̬�ļ�����
'*************************************************************************
Option Explicit

Dim St As Long

Private Sub Command1_Click()
                                '�����Ƕ�Text1�������໯����
If St = -1 Then

    PrevWndProc = SetWindowLong(Text1.Hwnd, GWL_WNDPROC, AddressOf SubWndProc)
    Command1.Caption = "ֹͣ"
    St = 1
    Me.Caption = "���ദ��״̬!"
     
Else

    SetWindowLong Text1.Hwnd, GWL_WNDPROC, PrevWndProc
    Command1.Caption = "��ʼ"
    St = -1
    Me.Caption = "����״̬"
    
End If

End Sub

Private Sub Form_Load()

St = -1
Me.Caption = "׼�����"

End Sub

Private Sub Form_Unload(Cancel As Integer)

If St <> -1 Then
    SetWindowLong Text1.Hwnd, GWL_WNDPROC, PrevWndProc
End If

End Sub

Private Sub Timer1_Timer()

Select Case MouseW
    Case 1
        Label1.Caption = "���Ϲ���"
    Case -1
        Label1.Caption = "���¹���"
    Case Else
        Label1.Caption = "���־�ֹ"
End Select

MouseW = 0

End Sub
VERSION 5.00
Begin VB.Form LvMath 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ߵ���ѧ"
   ClientHeight    =   6972
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10296
   Icon            =   "LvMath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6972
   ScaleWidth      =   10296
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7800
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label Bt 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "aBc"
      BeginProperty Font 
         Name            =   "����"
         Size            =   25.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   516
      Index           =   0
      Left            =   0
      MouseIcon       =   "LvMath.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   792
   End
   Begin VB.Label TiMu 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "aBc"
      BeginProperty Font 
         Name            =   "����"
         Size            =   25.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   516
      Index           =   0
      Left            =   4752
      TabIndex        =   0
      Top             =   120
      Width           =   792
   End
End
Attribute VB_Name = "LvMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tL As Integer
Const Sy As String = "* ʣ����Ŀ����"

Private Sub Form_Unload(Cancel As Integer)
Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
End
End Sub

Private Sub Bt_Click(Index As Integer)
Select Case tL
    Case 6
        Dui
        Bt(Index).BackColor = RGB(66, 255, 66)
        tL = 5
        Ld
        TiMu(2) = "2�� ��仰�м������֣�"
        Bt(2) = "8"
        Bt(3) = "14"
        TiMu(0) = Sy & "5"
    Case 5
        If Index = 2 Then
            Dui
            Bt(Index).BackColor = RGB(66, 255, 66)
            tL = 4
            Ld
            TiMu(3) = "3�� 3 + 4 ="
            Bt(4) = "6"
            Bt(5) = "7"
            TiMu(0) = Sy & "4"
        Else
            Si
        End If
    Case 4
        If Index = 5 Then
            Dui
            Bt(Index).BackColor = RGB(66, 255, 66)
            tL = 3
            Ld
            TiMu(4) = "4�� 7 - 1 ="
            Bt(6) = "6"
            Bt(7) = "99"
            TiMu(0) = Sy & "3"
        Else
            Si
        End If
    Case 3
        If Index = 6 Then
            Dui
            Bt(Index).BackColor = RGB(66, 255, 66)
            tL = 2
            Ld
            TiMu(5) = "5�� 1 X 5 ="
            Bt(8) = "1"
            Bt(9) = "5"
            TiMu(0) = Sy & "2"
        Else
            Si
        End If
    Case 2
        If Index = 9 Then
            Dui
            Bt(Index).BackColor = RGB(66, 255, 66)
            tL = 1
            Ld
            TiMu(6) = "6��A=[1,9),��DA."
            Bt(10).Visible = False
            Bt(11).Visible = False
            Text1.Visible = True
            TiMu(0) = Sy & "1"
        Else
            Si
        End If
End Select
End Sub

Private Sub Form_Load()
tL = 6
TiMu(0).Alignment = 2
TiMu(0) = Sy & tL
Ld
TiMu(1).FontSize = 22
TiMu(1) = "1��������ս��,�㽫���6����ѧ��," & Chr(10) & "���Ѷ�ѭ�򽥽�.������:һ������ѡ" & Chr(10) & "����,4��ѡ����,��һ�������."
With Bt(1)
    .FontSize = 22
    .Caption = "��������"
    .Top = .Top + 436
    .Left = 8006
End With
End Sub

Private Sub Ld()
Dim Ind As Integer
Ind = 7 - tL
Load TiMu(Ind)
With TiMu(Ind)
    .BackColor = 0
    .ForeColor = vbWhite
    .Alignment = 0
    .Top = TiMu(Ind - 1).Top + TiMu(Ind - 1).Height + 266
    .Left = 366
    .Visible = True
End With
If Ind <> 1 Then
    Load Bt(2 * Ind - 2)
    With Bt(2 * Ind - 2)
        .Top = TiMu(Ind).Top
        .Left = 7666
        .Visible = True
    End With
End If
Load Bt(2 * Ind - 1)
    With Bt(2 * Ind - 1)
        .Top = TiMu(Ind).Top
        .Left = 8806
        .Visible = True
    End With
End Sub

Private Sub Si()
Me.Visible = False
MsgBox "ϵͳ���������̸ж�֮��ĬĬ�ر����ˡ�", vbCritical + vbOKOnly, "ERROR"
MsgBox "����ȥð�ա��Ѿ��˳���"
End
End Sub

Private Sub Dui()
MsgBox "��ϲ�����ˡ�", vbOKOnly + vbInformation, "�ߵ���ѧ"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 1
        With Text1
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    Case 13
        If Text1 = "4" Then
            Win
        Else
            Si
        End If
End Select
End Sub

Private Sub Win()
SaveSetting "iGlope", "woqu", "j3", 1
MsgBox "��ͨ���˽����ؿ����ߵ���ѧ����", vbOKOnly + vbInformation, "��ϲ��"
Unload Me
End Sub

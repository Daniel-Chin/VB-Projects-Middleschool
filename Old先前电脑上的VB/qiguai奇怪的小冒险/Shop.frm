VERSION 5.00
Begin VB.Form Shop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�̵�"
   ClientHeight    =   6492
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10320
   Icon            =   "Shop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6492
   ScaleWidth      =   10320
   StartUpPosition =   2  '��Ļ����
   Begin VB.Label Bt 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   972
      Index           =   5
      Left            =   3600
      TabIndex        =   9
      Top             =   4560
      Width           =   3012
   End
   Begin VB.Label Bt 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   972
      Index           =   4
      Left            =   360
      TabIndex        =   8
      Top             =   4560
      Width           =   3012
   End
   Begin VB.Label Bt 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   972
      Index           =   3
      Left            =   6840
      TabIndex        =   7
      Top             =   2880
      Width           =   3012
   End
   Begin VB.Label Bt 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   972
      Index           =   2
      Left            =   3600
      TabIndex        =   6
      Top             =   2880
      Width           =   3012
   End
   Begin VB.Label Bt 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   6
      Left            =   7320
      TabIndex        =   5
      Top             =   5400
      Width           =   2532
   End
   Begin VB.Label Bt 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   972
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   3012
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   7200
      X2              =   10200
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   3000
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      Height          =   972
      Left            =   7800
      TabIndex        =   3
      Top             =   1320
      Width           =   2412
   End
   Begin VB.Label Label2 
      Height          =   972
      Left            =   7680
      TabIndex        =   2
      Top             =   360
      Width           =   2412
   End
   Begin VB.Label Label1 
      Height          =   972
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2412
   End
   Begin VB.Label BiaoTi 
      Caption         =   "�̵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   4074
      TabIndex        =   0
      Top             =   240
      Width           =   2172
   End
End
Attribute VB_Name = "Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cost As Integer
Dim NowOn As Byte
Dim Nm As Boolean

Private Sub Bt_Click(Index As Integer)
Select Case Index
    Case 1
        Cost = Cost + 3
    Case 2
        Cost = Cost + 18
    Case 3
        Cost = Cost + 21
    Case 4
        Cost = Cost + 199
    Case 6
        If MsgBox("�����ȷ��Ҫ������Щ��Ʒ���뵥����ȡ������������ȷ�����Է����̵ꡣ", vbQuestion + vbOKCancel, "ȷ�Ϲ���") = vbCancel Then
            MsgBox "Just for fun! ��ֻ��һ���ʵ����벻���Mojang��ƷMinecraft�¾�������", vbOKOnly, "�������һ���ʵ�����"
            Nm = True
            MainMenu.Visible = True
            Unload Me
        End If
End Select
Label3 = Cost
End Sub

Private Sub Bt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove 0, 0, 0, 0
With Bt(Index)
    .BackColor = vbWhite
    .ForeColor = 0
End With
NowOn = Index
End Sub

Private Sub Form_Load()
Label1.FontSize = 20
Label2.FontSize = 20
Label3.FontSize = 30
Label1 = "������ӵ�У�" & Chr(13) & Chr(10) & " 20������ֵ��"
Label2 = "��ι��򽫻�������ֵ��"
Label3 = "0"
For I = 1 To 6
    Bt(I).FontSize = 20
Next I
Bt(5).FontSize = 30
Bt(1) = "һ��Jimp" & Chr(13) & Chr(10) & "�۸�3������"
Bt(2) = "���Jimp" & Chr(13) & Chr(10) & "�۸�18������"
Bt(3) = "����ҩ��" & Chr(13) & Chr(10) & "�۸�21������"
Bt(4) = "�ؼ��б�" & Chr(13) & Chr(10) & "�۸�199������"
Bt(5) = "��δ������"
Bt(6) = "���˲���Ǯ"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        For I = 1 To 6
            With Bt(I)
                If I = NowOn Then GoTo Skip
                .BackColor = 0
                .ForeColor = vbWhite
            End With
Skip:   Next I
        NowOn = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not Nm Then
    End
End If
End Sub

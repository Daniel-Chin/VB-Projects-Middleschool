VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��°ͺղ���"
   ClientHeight    =   5484
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   5484
   ScaleWidth      =   9240
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Main 
      Left            =   5640
      Top             =   2520
   End
   Begin VB.Label Starter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "��ʼ"
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
      Left            =   3414
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2412
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SuShuTop As Long = 299999
Dim S(1 To 299999) As Boolean
Dim t As Long
Dim oS(1 To 999999) As Long, Ge(1 To 999999, 0 To 1) As Long, sU(1 To 999999) As Long, cHa(1 To 999999) As Integer
Dim DengChang(1 To 299999) As Boolean


Private Function IfSu(Number As Long) As Boolean
IfSu = True
For k = 2 To Int(Sqr(Number))
    If S(k) = True Then
        If (Number \ k) * k = Number Then
            IfSu = False
            Exit For
        End If
    End If
Next k
End Function

Private Sub Form_Load()
On Error Resume Next
Kill "F:\MatLab\GeDe\1.txt"
Open "F:\MatLab\GeDe\1.txt" For Output As #1
Starter.Visible = False
Show
If MsgBox("��������" & SuShuTop & "���µ�ȫ������������OK������Cancel�˳���", vbOKCancel + vbInformation, "��°ͺղ���") = vbCancel Then
    End
End If
Form1.Caption = "��������" & SuShuTop & "���µ�ȫ������"
Dim k As Long
For k = 2 To SuShuTop
    S(k) = IfSu(k)
    If k Mod 10000 = 0 Then Form1.Caption = Form1.Caption & "��"
Next k
Form1.Caption = "��°ͺղ���"
MsgBox "������ϡ�", vbInformation + vbOKOnly, "��°ͺղ���"
Starter.Visible = True
End Sub

Private Sub Main_Timer()
For kk = 0 To 63
Print #1, t; oS(t); Ge(t, 0); Ge(t, 1); sU(t) & ";..."
t = t + 1
oS(t) = 2 * t
sU(t) = -t * S(t)
For k = 2 * t To 2 Step -1
    If DengChang(k) = True Then
        If DengChang(2 * t - k) = True Then
            GoTo JieShu
        End If
    End If
Next k
For k = t To t * 2
    If S(k) = True And S(t * 2 - k) = True Then
        If Not DengChang(t * 2 - k) Then
            Ge(t, 0) = t * 2 - k
            DengChang(t * 2 - k) = True
        End If
        If Not DengChang(k) Then
            Ge(t, 1) = k
            DengChang(k) = True
        End If
        GoTo JieShu
    End If
Next k
MsgBox "��°ͺղ��뷴����" & t * 2, vbCritical + vbOKOnly, "FUCK"
JieShu:
Next kk
End Sub

Private Sub Starter_Click()
If Starter = "�ر�" Then
    Close #1
    End
Else
    Main.Interval = 1
    t = 1
    Starter = "�ر�"
    Me.Caption = "��������"
End If
End Sub

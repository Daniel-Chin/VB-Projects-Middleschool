VERSION 5.00
Begin VB.Form Level1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ֵ�Сð�� - ��1��"
   ClientHeight    =   4752
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10620
   Icon            =   "Level1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4752
   ScaleWidth      =   10620
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Main 
      Interval        =   30
      Left            =   8160
      Top             =   600
   End
   Begin VB.Label Ground 
      BackColor       =   &H00000000&
      Caption         =   "������ƶ�����Ծ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1692
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   7812
   End
   Begin VB.Shape I 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   612
      Left            =   1320
      Top             =   960
      Width           =   852
   End
   Begin VB.Label Win 
      BackColor       =   &H0000FFFF&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   852
      Left            =   6960
      TabIndex        =   1
      Top             =   480
      Width           =   852
   End
End
Attribute VB_Name = "Level1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Speed As Byte = 100
Dim VerSpeed As Integer
Dim Nm As Boolean
Dim z As Boolean, y As Boolean, s As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37     '��
        z = True
    Case 38     '��
        s = True
    Case 39    '��
        y = True
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37     '��
        z = False
    Case 38     '��
        s = False
    Case 39    '��
        y = False
End Select
End Sub

Private Sub Form_Load()
Nm = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not Nm Then
    End
End If
End Sub

Private Sub Main_Timer()
I.Top = I.Top - VerSpeed
If I.Top + I.Height < Ground.Top Then
    VerSpeed = VerSpeed - 5
Else
    VerSpeed = 0
End If
If I.Left + I.Width > Win.Left And I.Left < Win.Left + Win.Width And I.Top < Win.Top + Win.Height And I.Top + I.Height > Win.Top Then
    MsgBox "��Ӯ�ˣ����ڽ���ڶ��أ�", vbOKOnly + vbInformation, "��ϲ"
    Level2.Visible = True
    Nm = True
    z = False
    y = False
    s = False
    Unload Me
    Exit Sub
End If

    If z Then
        If I.Left > 0 Then
            I.Left = I.Left - Speed
        End If
    End If
    If s Then
        If I.Top + I.Height >= Ground.Top Then
            VerSpeed = 100
        End If
    End If
    If y Then
        If I.Left + I.Width < Me.Width Then
            I.Left = I.Left + Speed
        End If
    End If
End Sub

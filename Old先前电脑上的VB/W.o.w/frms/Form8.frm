VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "ð��ģʽ"
   ClientHeight    =   7536
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8652
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7536
   ScaleWidth      =   8652
   Begin VB.Timer Timer2 
      Left            =   6360
      Top             =   4800
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0000FF00&
      Caption         =   "Command9"
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
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   4092
   End
   Begin VB.TextBox Text16 
      Height          =   372
      Left            =   6480
      TabIndex        =   31
      Top             =   4200
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��Ҫʹ���ؼ�"
      Height          =   492
      Left            =   6480
      TabIndex        =   30
      Top             =   3480
      Width           =   1812
   End
   Begin VB.CommandButton Command7 
      Caption         =   "��ɱ"
      Height          =   372
      Left            =   6480
      TabIndex        =   29
      Top             =   3000
      Width           =   1812
   End
   Begin VB.TextBox Text15 
      Height          =   492
      Left            =   9000
      TabIndex        =   27
      Text            =   "0"
      Top             =   720
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Option"
      Height          =   492
      Left            =   6480
      TabIndex        =   26
      Top             =   2400
      Width           =   1812
   End
   Begin VB.TextBox Text12 
      Height          =   492
      Left            =   9960
      TabIndex        =   25
      Text            =   "0"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.TextBox Text11 
      Height          =   372
      Left            =   9960
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.ComboBox Combo2 
      Height          =   276
      ItemData        =   "Form8.frx":0000
      Left            =   7560
      List            =   "Form8.frx":000A
      TabIndex        =   19
      Text            =   "����"
      Top             =   2040
      Width           =   972
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "Form8.frx":001A
      Left            =   6480
      List            =   "Form8.frx":002A
      TabIndex        =   18
      Text            =   "�Ѷ�"
      Top             =   2040
      Width           =   852
   End
   Begin VB.TextBox Text14 
      Height          =   264
      Left            =   10560
      TabIndex        =   15
      Text            =   "0"
      Top             =   3240
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox Text13 
      Height          =   372
      Left            =   7560
      TabIndex        =   14
      Text            =   "12"
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Height          =   372
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   372
   End
   Begin VB.TextBox Text9 
      Height          =   372
      Left            =   12840
      TabIndex        =   12
      Text            =   "-25"
      Top             =   4320
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Height          =   372
      Left            =   12600
      TabIndex        =   11
      Text            =   "25"
      Top             =   3600
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text7 
      Height          =   372
      Left            =   12600
      TabIndex        =   10
      Text            =   "30"
      Top             =   3000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text6 
      Height          =   372
      Left            =   12600
      TabIndex        =   9
      Text            =   "28"
      Top             =   2520
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text5 
      Height          =   372
      Left            =   12960
      TabIndex        =   8
      Text            =   "20"
      Top             =   1800
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text4 
      Height          =   372
      Left            =   12840
      TabIndex        =   7
      Text            =   "-30"
      Top             =   1200
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF0000&
      Height          =   492
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   732
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Height          =   492
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   612
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Height          =   660
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   372
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Height          =   372
      Left            =   2040
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1452
   End
   Begin VB.TextBox Text3 
      Height          =   1212
      Left            =   9600
      TabIndex        =   2
      Text            =   "0"
      Top             =   2160
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   12840
      TabIndex        =   1
      Text            =   "-20"
      Top             =   600
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   13200
      TabIndex        =   0
      Text            =   "-25"
      Top             =   120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Timer Timer1 
      Left            =   11160
      Top             =   6240
   End
   Begin VB.Label Label17 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ǯ����Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   25.8
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4000
      Left            =   120
      TabIndex        =   42
      Top             =   360
      Width           =   612
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00008080&
      BorderWidth     =   5
      Height          =   4452
      Left            =   120
      Top             =   360
      Width           =   612
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   240
      TabIndex        =   41
      Top             =   6360
      Width           =   852
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   8
      Height          =   372
      Left            =   120
      Top             =   6360
      Width           =   6012
   End
   Begin VB.Label Label15 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   372
      Left            =   120
      TabIndex        =   40
      Top             =   6360
      Width           =   6012
   End
   Begin VB.Label Label14 
      Caption         =   "1"
      Height          =   252
      Left            =   7560
      TabIndex        =   39
      Top             =   1680
      Width           =   612
   End
   Begin VB.Label Label13 
      Caption         =   "�ȼ�"
      Height          =   252
      Left            =   6480
      TabIndex        =   38
      Top             =   1680
      Width           =   852
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      Height          =   252
      Left            =   7560
      TabIndex        =   37
      Top             =   1320
      Width           =   1092
   End
   Begin VB.Label Label11 
      Caption         =   "���飺"
      Height          =   252
      Left            =   6480
      TabIndex        =   36
      Top             =   1320
      Width           =   852
   End
   Begin VB.Label Label10 
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   7560
      TabIndex        =   35
      Top             =   960
      Width           =   1092
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   252
      Left            =   7560
      TabIndex        =   34
      Top             =   600
      Width           =   852
   End
   Begin VB.Label Label8 
      Caption         =   "��Ǯ��"
      Height          =   252
      Left            =   6480
      TabIndex        =   33
      Top             =   960
      Width           =   852
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   252
   End
   Begin VB.Label Label6 
      Caption         =   "��߷֣�"
      Height          =   252
      Left            =   6480
      TabIndex        =   23
      Top             =   600
      Width           =   852
   End
   Begin VB.Label Label5 
      Caption         =   "�����ǵ�1��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   960
      TabIndex        =   22
      Top             =   600
      Width           =   1932
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1812
      Left            =   720
      TabIndex        =   21
      Top             =   1080
      Width           =   2532
   End
   Begin VB.Label Label2 
      Caption         =   "��Ļ���أ�"
      Height          =   372
      Left            =   6480
      TabIndex        =   20
      Top             =   120
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "time"
      Height          =   612
      Left            =   9120
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "Check"
      Height          =   372
      Left            =   9720
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   322
      Left            =   0
      Top             =   3172
      Width           =   3852
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   3492
      Left            =   3532
      Top             =   0
      Width           =   372
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lala_la As Boolean
Dim yinlema As Boolean    '����ʱ�ж�
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim position As POINTAPI
Public money As Long     '��Ǯ
Public money_mt As Double   '��Ǯ����
Public exp As Long    '����
Public level As Integer '�ȼ�
Dim achi_load As Long
Public achi_combo As Boolean    '��������
Public achi_achibtm As Boolean   '����ɾ��»���ť
Public achi_menmen As Boolean   '����
Public achi_win As Boolean    '�ӳ�����
Public achi_start As Boolean     '��ʼ��ʼ
Public achi_blood As Boolean     '��һ��Ѫ
Public achi_shenji As Boolean     '������
Public achi_option As Boolean     '��Ҫ����
Public achi_mijimiji As Boolean    '�ؼ��ؼ�
Public achi_cap As String
Dim nmal As Boolean
Private Sub Form_Unload(Cancel As Integer)
If nmal = True Then
Else
    End
End If
End Sub

Private Sub ��ʼ�ĳɾ�()
    If achi_start = False Then    '�ɾ͡���ʼ��
        achi_start = True
        achi_cap = "��ʼ����ʼ��"
        Call achibb1
    End If
End Sub

Private Sub Command5_Click()
Combo1.Visible = False
Combo2.Visible = False
If Int(exp / 100 + 1) <= 100 Then     '���µȼ�
    level = Int(exp / 100 + 1)
Else
    level = 100
End If
If level > 1 And achi_shenji = False Then     '�ɾͣ�������
    achi_shenji = True       '�˳ɾ�����ˣ�
    achi_cap = "������"    '��ʾ����
    Call achibb1
End If
Select Case level    '��Ǯ����
    Case 1
        money_mt = 0.01
    Case 2
        money_mt = 0.03
    Case 3
        money_mt = 0.05
    Case 4
        money_mt = 0.09
    Case 4 To 20
        money_mt = level / 45
    Case Is > 20
        money_mt = level / 10
End Select
money_mt = level / 100
If Combo1.Text = "�Ѷ�" Then
    MsgBox "��ѡ����Ϸ�Ѷȣ�", vbExclamation, "����"
    Combo1.Visible = True
    Combo2.Visible = True
Else
    If Text14 = 0 Then
        Timer1.Interval = 50
    End If
    Select Case Combo1.Text
        Case "��"
            ��ʼ�ĳɾ�
            Shape2.Top = 6000
            Shape1.Left = 6000
            Shape1.Height = 6000
            Shape2.Width = 6000
        Case "��ͨ"
            ��ʼ�ĳɾ�
            Shape2.Top = 4000
            Shape1.Left = 4000
            Shape1.Height = 4000
            Shape2.Width = 4000
        Case "����"
            ��ʼ�ĳɾ�
            Shape2.Top = 3000
            Shape1.Left = 3000
            Shape1.Height = 3000
            Shape2.Width = 3000
            Command1.Top = 1
            Command1.Left = 1
        Case "����"
            ��ʼ�ĳɾ�
            Shape2.Top = 2000
            Shape1.Left = 2000
            Shape1.Height = 2000
            Shape2.Width = 2000
            Command1.Visible = False
            Command2.Visible = False
            Command4.Visible = False
        Case "��������"
            Timer1.Interval = 0        '��Ҫ��ʼ
            If achi_combo = False Then
                achi_combo = True       '�˳ɾ�����ˣ�
                achi_cap = "�治������"    '��ʾ����
                Call achibb1              '���С��ɾ��»���
            End If
        Case Else
            Timer1.Interval = 0         '��Ҫ��ʼ
            Shape2.Top = 6000
            Shape1.Left = 6000
            Shape1.Height = 6000
            Shape2.Width = 6000
            MsgBox "��Ҫ������ѡ�������������������", 1, "�ⲻ��һ���ʵ����ҷ�4"
            Combo1.Visible = True
            Combo2.Visible = True
    End Select
End If
End Sub
Private Sub achibb1()    'form1�ĳɾ��»���bb��bubble����˼
Command9.Caption = "��óɾͣ�" & achi_cap
achi_load = -40
Timer2.Interval = 30
End Sub
Private Sub Command6_Click()
    Form3.Visible = True
    If achi_option = False Then        '�ɾͣ���Ҫ����
        achi_option = True
        achi_cap = "��Ҫ����"
        Call achibb1
    End If
End Sub

Private Sub Command7_Click()
Timer1.Interval = 0
Call ����
Combo2.Text = "����"
End Sub

Private Sub Command8_Click()
If Text16.Visible = True Then Text16.Visible = False Else Text16.Visible = True: Text16.SetFocus
End Sub

Private Sub Command9_Click()
If achi_achibtm = False Then         '�ɾͣ���ť
    achi_achibtm = True
    achi_cap = "������ť�����"
    Call achibb1
End If
End Sub

Private Sub Form_Load()
If lala_la <> True Then
    Call ����
    Call MsgBox("�������ܼ�ּ��룡����", vbExclamation, "˵��")
    nmal = False
End If
End Sub

Private Sub ����()
Label7.Visible = True
If Command5.Top < Label7.Height And Command5.Left < Label7.Width Then
    Label4 = "��Ӯ�ˣ�": Timer1.Interval = 0: Call MsgBox("��ϲ����ͨ���ˣ�������Ϸģʽ�����ˣ�", vbExclamation, "ͨ��"): Text15 = 1: Call ����
    yinlema = True
    If achi_win = False Then     '�ɾͣ��ӳ�����
        achi_win = True
        achi_cap = "�ӳ����죡"
        Call achibb1
    End If
End If
End Sub

Private Sub Text16_Change()
Select Case Text16
    Case "Wendy"
        Call �ؼ��ɾ�
        Call ����
    Case "��̫����"
        Call �ؼ��ɾ�
        mnvalue = InputBox("��������Ҫ�Ľ������", "���ײ��Ǻõ���Ϊ", "0")
        money = money + mnvalue
    'Case                                                                                           δ�����
End Select
End Sub
Private Sub �ؼ��ɾ�()            '�ɾͣ��ؼ��ؼ�
If achi_mijimiji = False Then
    achi_mijimiji = True
    achi_cap = "�ؼ��ؼ�"
    Call achibb1
End If
End Sub
Private Sub Timer1_Timer()
Label12 = exp
Label14 = level '������Ѿ�����ʾ����
exp1 = exp                 '������Ϊ���ıȶԾ���ֵ��Ǯ�仯���̵�
money1 = money
exp = exp + Int(Rnd() * 1.05)
money = money + Int(Rnd() * (1 + money_mt))  '���������Ǽ�Ǯ
Label10 = money
Text3 = Text3 + 1           '��ʱ
If exp <> exp1 Then
    Label15.Width = (exp Mod 100) * 6012 / 100    '���������ȱ仯������6012�ǳ�ʼ����
End If
If money > money1 Then
    If Label17.Height < 4452 Then          'Ǯ�����ȱ仯������4452�ǳ�ʼ����
        Label17.Height = Label17.Height + 4452 / 100 * (money - money1)
    Else
        Label17.Height = 1             '����
        exp = exp + 10            '�Ӿ���
        Label15.Width = (exp Mod 100) * 6012 / 100          '���¾�����
    End If
End If
If Text12 = 1 Then Call ����
If Command1.Top < 0 Or Command1.Top > Shape2.Top - Command1.Height Then Text1 = -Text1
If Command1.Left < 0 Or Command1.Left > Shape1.Left - Command1.Width Then Text2 = -Text2
If Command2.Top < 0 Or Command2.Top > Shape2.Top - Command2.Height Then Text4 = -Text4
If Command2.Left < 0 Or Command2.Left > Shape1.Left - Command2.Width Then Text5 = -Text5
If Command3.Top < 0 Or Command3.Top > Shape2.Top - Command3.Height Then Text6 = -Text6
If Command3.Left < 0 Or Command3.Left > Shape1.Left - Command3.Width Then Text7 = -Text7
If Command4.Top < 0 Or Command4.Top > Shape2.Top - Command4.Height Then Text8 = -Text8
If Command4.Left < 0 Or Command4.Left > Shape1.Left - Command4.Width Then Text9 = -Text9
Command1.Top = Command1.Top + Text1
Command1.Left = Command1.Left + Text2
Command2.Top = Command2.Top + Text4
Command2.Left = Command2.Left + Text5
Command3.Top = Command3.Top + Text6
Command3.Left = Command3.Left + Text7
Command4.Top = Command4.Top + Text8
Command4.Left = Command4.Left + Text9
If Text3 < 5 Then Timer1.Interval = 100: Label4 = "�������������������˿�ʼ׷���ˣ�": Label5 = "�����ǵ�1�ء�"
If Text3.Text > 5 And Text3 < 8 Then Timer1.Interval = 80
If Text3.Text > 8 And Text3 < 10 Then Timer1.Interval = 60
If Text3.Text > 10 And Text3 < 25 Then Timer1.Interval = 40
If Text3.Text > 25 And Text3 < 100 Then Timer1.Interval = 30: Label4 = "���ã������в��Ǻܿ졣���������ڿ�ʼ�����ǽ����úܿ죡�������ָ�ӹٱ�ͷ��У��": Label5 = " �����ǵ�2�ء�": Text11 = 2
If Text3.Text > 100 And Text3 < 300 Then Timer1.Interval = 20:
If Text3.Text > 300 And Text3 < 1000 Then Timer1.Interval = 15: Label4 = "���Ѿ�Ϊ�����ϼ�����������ˡ��ټ��һ�����": Label5 = "�����ǵ�3�ء�": Text11 = 3
If Text3.Text > 1000 And Text3 < 1100 Then Timer1.Interval = 60: Label4 = "���������ˣ�Ŷ������������������ţ��": Label5 = "�����ǵ�4�ء�": Text11 = 4 '500��600�Ǳ�����
If Text3.Text > 1100 And Text3 < 2000 Then Timer1.Interval = 15: Label4 = "����Ч�������ˡ��ټ��һ������������ϾͿ���": Label5 = "�����ǵ�5�ء�": Text11 = 5 '��2000�Ŵ�
If Text3.Text > 2000 Then Label4 = "�ſ��ˣ����Ͻǣ��������ܰɣ�": Text12 = 1: Label5 = "�����ǵ�6�ء�": Text11 = 6
GetCursorPos position
ScreenToClient Me.hwnd, position
Command5.Top = position.y * Text13 - 150
Command5.Left = position.x * Text13 - 150
Call ����
If Text14 = 1 And Combo2.Text = "����" And yinlema = False Then
        If achi_menmen = False And Text3 > 2000 Then       '�ɾͣ�������
            achi_menmen = True
            achi_cap = "�ţ��ţ�"
            Call achibb1
        End If
    Timer1.Interval = 0
    If Text3 * 1 > Label9 * 1 Then Label9 = Text3
    Label4 = "�����ˣ������������¿�ʼ"
    Call MsgBox("�ǲ�����һ�ֱ�ƭ�ĸо���", vbExclamation, "��£�")
    nmal = True
    Unload Form8
    Form9.Visible = True
    Timer2.Interval = 0
End If
End Sub
Sub ����()
If Command5.Top > (Command1.Top - Command5.Height) And Command5.Top < (Command1.Top + Command1.Height) And Command5.Left < (Command1.Left + Command1.Width) And Command5.Left > (Command1.Left - Command5.Width) Then Text14 = 1
If Command5.Top > (Command2.Top - Command5.Height) And Command5.Top < (Command2.Top + Command2.Height) And Command5.Left < (Command2.Left + Command2.Width) And Command5.Left > (Command2.Left - Command5.Width) Then Text14 = 1
If Command5.Top > (Command3.Top - Command5.Height) And Command5.Top < (Command3.Top + Command3.Height) And Command5.Left < (Command3.Left + Command3.Width) And Command5.Left > (Command3.Left - Command5.Width) Then Text14 = 1
If Command5.Top > (Command4.Top - Command5.Height) And Command5.Top < (Command4.Top + Command4.Height) And Command5.Left < (Command4.Left + Command4.Width) And Command5.Left > (Command4.Left - Command5.Width) Then Text14 = 1
If Command5.Top > Shape2.Top - Command5.Height Or Command5.Left > Shape1.Left - Command5.Width Or Command5.Top < 0 Or Command5.Left < 0 Then Text14 = 1
End Sub
Sub ����()
Text11 = 1
Text12 = 0
Label7.Visible = False
Combo1.Visible = True
Combo2.Visible = True
Command1.Top = 2280
Command1.Left = 2040
Command1.Visible = True
Command5.Top = 1320
Command5.Left = 1200
Command5.Visible = True
Command2.Top = 2280
Command2.Left = 120
Command2.Visible = True
Command3.Top = 120
Command3.Left = 120
Command3.Visible = True
Command4.Top = 0
Command4.Left = 2040
Command4.Visible = True
Text14 = 0
Text3 = 0
yinlema = False
Label17.Height = 1         '���Ǯ����Ϣ
End Sub

Private Sub ����()
                  '��������
End Sub

Private Sub Timer2_Timer()
Command9.Top = 40 - achi_load ^ 2
Command9.Visible = True
achi_load = achi_load + 1
If achi_load > 50 Then Timer2.Interval = 0: Command9.Visible = False
End Sub

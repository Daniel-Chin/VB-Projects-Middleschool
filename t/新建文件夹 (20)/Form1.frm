VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6024
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   10848
   LinkTopic       =   "Form1"
   ScaleHeight     =   6024
   ScaleWidth      =   10848
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Main 
      Interval        =   40
      Left            =   4920
      Top             =   2880
   End
   Begin VB.Shape Wall 
      FillColor       =   &H00C0C000&
      FillStyle       =   6  'Cross
      Height          =   732
      Index           =   6
      Left            =   3600
      Top             =   1920
      Width           =   732
   End
   Begin VB.Shape Wall 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   6  'Cross
      Height          =   372
      Index           =   4
      Left            =   7680
      Top             =   1920
      Width           =   3492
   End
   Begin VB.Shape Wall 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   6  'Cross
      Height          =   972
      Index           =   3
      Left            =   6480
      Top             =   600
      Width           =   252
   End
   Begin VB.Shape Wall 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   6  'Cross
      Height          =   1092
      Index           =   2
      Left            =   6240
      Top             =   3000
      Width           =   4212
   End
   Begin VB.Shape Wall 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   6  'Cross
      Height          =   1092
      Index           =   1
      Left            =   5640
      Top             =   4560
      Width           =   1212
   End
   Begin VB.Shape Wall 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   6  'Cross
      Height          =   1092
      Index           =   5
      Left            =   1200
      Top             =   3960
      Width           =   4212
   End
   Begin VB.Shape Player 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   852
      Left            =   2400
      Top             =   2160
      Width           =   492
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const G As Long = 10 '�������ٶ�
Const Walk_V As Long = 50 '�����ƶ��ٶ�
Const Jump_V As Long = 200 '��Ծ���ٶ�

Dim Vert_Speed As Long '��ֱ�ٶ�

'��װ�������Ұ���
Private Type Keys
    Zuo As Boolean
    You As Boolean
    Jump As Boolean
End Type

Dim WASD As Keys

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 38
        WASD.Jump = True
    Case 37
        WASD.Zuo = True
    Case 39
        WASD.You = True
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 38
        WASD.Jump = False
    Case 37
        WASD.Zuo = False
    Case 39
        WASD.You = False
End Select
End Sub
'��װ��β�������Ұ���

Private Function Qiang(Optional XLeft As Long = -666, Optional XTop As Long = -666) As Byte
'����������������Ϊ(Xleft,XTop)������
'��ǽ�Ƿ��ص�������Qiang����0��
'���ǣ�Qiang�������ص���ǽ��Index��

If XLeft = -666 Then XLeft = Player.left
If XTop = -666 Then XTop = Player.top
'���ʡ���Ա��������������Զ�ȡ���ǵ�ǰ����ʵ���ꡣ

Dim A(1 To 4) As Boolean '��ʱ����
Dim k As Byte
For k = 1 To Wall.UBound '�������е�ǽ
    With Wall(k)
        A(1) = XLeft + Player.Width >= .left
        A(2) = XLeft <= .left + .Width
        A(3) = XTop + Player.Height >= .top
        A(4) = XTop <= .top + .Height
        If A(1) And A(2) And A(3) And A(4) Then
            '���ص�
            Qiang = k
            Exit Function
        End If
    End With
Next k

'û���ص�
Qiang = 0 '���Ҳ����ʡ��
End Function

Private Function isGrnd() As Boolean
'���ǵ�ǰ�Ƿ��ŵأ�
If Qiang(, Player.top + 1) Then isGrnd = True
End Function

Private Sub Main_Timer()
Dim B As Long '��ʱ�����������ݸ塱�Ĺ���
With Player
    '����λ��
    B = .left + (WASD.Zuo - WASD.You) * Walk_V '����ر�����
    'B��¼���ǵġ�Ը��λ�á��ĺ�����
    Do While Qiang(B)
        '���ѭ���Ĺ��ܾ��ǣ����Ը��λ����ǽ��
        '������Ը��λ�ã�ֱ��ûǽΪֹ��
        B = B + WASD.You - WASD.Zuo
    Loop
    .left = B '��ɺ���Ը����
    
    '����λ��
    Dim Grnd As Boolean '�Ƿ��ŵ�
    Grnd = isGrnd()
    If Grnd Then
        '�ŵ�
        If WASD.Jump Then '�����Ұ�ס����Ծ��
            '��Ծ
            Vert_Speed = -Jump_V
        Else
            '�����ٶȹ���(��Ϊ�ŵ����)
            Vert_Speed = 0
        End If
    Else
        '�ڿ�
        Vert_Speed = Vert_Speed + G
        '�������ٶ����ٶȵı仯��
    End If
    B = .top + Vert_Speed '��B��¼��Ը��λ�á���������
    If Vert_Speed < 0 Then '�����Ϸ���
        '�����Ϸ�
        If Qiang(, B) Then
            '�����ǽ
            Vert_Speed = 0 'ײ����ͷ�������ٶȹ���
            Do While Qiang(, B)
                '���ѭ���Ĺ��ܾ��ǣ����Ը��λ����ǽ��
                '������Ը��λ�ã�ֱ��ûǽΪֹ��
                B = B + 1
            Loop
            '��������൱�ڡ�����ǽ����һ�¼���
            '������ǽ��ID��Qiang(, B-1)
            '��ʲô�����ʺš��Ĺ��ܾ�������ʵ��
            
            If Qiang(, B - 1) = 6 Then '���Ǹ�����
                '�������6��ǽ
                Me.BackColor = Rnd * 256 '�ı���ɫ
            End If
            '''
        End If
    Else
        '�����µ�
        If Qiang(, B) Or Grnd Then
            '�ŵ�
            Vert_Speed = 0 '�����ٶȹ���
            Do While Qiang(, B)
                '���ѭ���Ĺ��ܾ��ǣ����Ը��λ����ǽ��
                '������Ը��λ�ã�ֱ��ûǽΪֹ��
                B = B - 1
            Loop
            '��������൱�ڡ��ȵ�ǽ����һ�¼���
            '������ǽ��ID��Qiang(, B+1)
            '��ʲô�����ʺš��Ĺ��ܾ�������ʵ��
            
            If Qiang(, B + 1) = 6 Then '���Ǹ�����
                '����ȵ�6��ǽ
                Wall(6).top = Wall(6).top + 200 '6��ǽ�½�
            End If
            '''
        End If
    End If
    .top = B '�������Ը����
End With
End Sub

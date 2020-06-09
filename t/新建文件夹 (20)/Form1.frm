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
   StartUpPosition =   3  '窗口缺省
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

Const G As Long = 10 '重力加速度
Const Walk_V As Long = 50 '横向移动速度
Const Jump_V As Long = 200 '跳跃初速度

Dim Vert_Speed As Long '垂直速度

'封装：检测玩家按键
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
'封装结尾：检测玩家按键

Private Function Qiang(Optional XLeft As Long = -666, Optional XTop As Long = -666) As Byte
'　　假设主角坐标为(Xleft,XTop)，主角
'和墙是否重叠？若否，Qiang返回0。
'若是，Qiang返回所重叠的墙的Index。

If XLeft = -666 Then XLeft = Player.left
If XTop = -666 Then XTop = Player.top
'如果省略自变量，假设坐标自动取主角当前的真实坐标。

Dim A(1 To 4) As Boolean '临时变量
Dim k As Byte
For k = 1 To Wall.UBound '遍历所有的墙
    With Wall(k)
        A(1) = XLeft + Player.Width >= .left
        A(2) = XLeft <= .left + .Width
        A(3) = XTop + Player.Height >= .top
        A(4) = XTop <= .top + .Height
        If A(1) And A(2) And A(3) And A(4) Then
            '有重叠
            Qiang = k
            Exit Function
        End If
    End With
Next k

'没有重叠
Qiang = 0 '这句也可以省略
End Function

Private Function isGrnd() As Boolean
'主角当前是否着地？
If Qiang(, Player.top + 1) Then isGrnd = True
End Function

Private Sub Main_Timer()
Dim B As Long '临时变量。作“草稿”的功能
With Player
    '横向位移
    B = .left + (WASD.Zuo - WASD.You) * Walk_V '这句特别巧妙
    'B记录主角的“愿望位置”的横坐标
    Do While Qiang(B)
        '这个循环的功能就是，如果愿望位置有墙，
        '就削减愿望位置，直到没墙为止。
        B = B + WASD.You - WASD.Zuo
    Loop
    .left = B '达成横向愿望。
    
    '纵向位移
    Dim Grnd As Boolean '是否着地
    Grnd = isGrnd()
    If Grnd Then
        '着地
        If WASD.Jump Then '如果玩家按住了跳跃键
            '跳跃
            Vert_Speed = -Jump_V
        Else
            '纵向速度归零(因为着地了嘛。)
            Vert_Speed = 0
        End If
    Else
        '腾空
        Vert_Speed = Vert_Speed + G
        '重力加速度是速度的变化量
    End If
    B = .top + Vert_Speed '用B记录“愿望位置”的纵坐标
    If Vert_Speed < 0 Then '在向上飞吗？
        '在向上飞
        If Qiang(, B) Then
            '如果有墙
            Vert_Speed = 0 '撞到了头！纵向速度归零
            Do While Qiang(, B)
                '这个循环的功能就是，如果愿望位置有墙，
                '就削减愿望位置，直到没墙为止。
                B = B + 1
            Loop
            '下面空行相当于“顶到墙”这一事件。
            '顶到的墙的ID即Qiang(, B-1)
            '像什么“顶问号”的功能就在这里实现
            
            If Qiang(, B - 1) = 6 Then '这是个范例
                '如果顶到6号墙
                Me.BackColor = Rnd * 256 '改背景色
            End If
            '''
        End If
    Else
        '在向下掉
        If Qiang(, B) Or Grnd Then
            '着地
            Vert_Speed = 0 '纵向速度归零
            Do While Qiang(, B)
                '这个循环的功能就是，如果愿望位置有墙，
                '就削减愿望位置，直到没墙为止。
                B = B - 1
            Loop
            '下面空行相当于“踩到墙”这一事件。
            '顶到的墙的ID即Qiang(, B+1)
            '像什么“踩问号”的功能就在这里实现
            
            If Qiang(, B + 1) = 6 Then '这是个范例
                '如果踩到6号墙
                Wall(6).top = Wall(6).top + 200 '6号墙下降
            End If
            '''
        End If
    End If
    .top = B '达成纵向愿望。
End With
End Sub

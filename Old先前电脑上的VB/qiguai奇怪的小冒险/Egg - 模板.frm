VERSION 5.00
Begin VB.Form Egg99 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "真正奇怪的冒险"
   ClientHeight    =   6336
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10260
   Icon            =   "Egg2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6336
   ScaleWidth      =   10260
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Focuser 
      Height          =   372
      Left            =   4680
      TabIndex        =   1
      Top             =   -500
      Width           =   972
   End
   Begin 奇怪的小冒险.LinXin Win 
      Height          =   1000
      Left            =   9200
      TabIndex        =   0
      Top             =   4000
      Width           =   800
      _extentx        =   1418
      _extenty        =   1757
   End
   Begin VB.Timer Main 
      Left            =   2520
      Top             =   2040
   End
   Begin VB.Timer Shaker 
      Interval        =   30
      Left            =   4560
      Top             =   3960
   End
   Begin VB.Label Q 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   700
      Left            =   3500
      TabIndex        =   2
      Top             =   600
      Width           =   700
   End
   Begin VB.Line W 
      BorderWidth     =   5
      Index           =   11
      X1              =   9216
      X2              =   9216
      Y1              =   3000
      Y2              =   3300
   End
   Begin VB.Line W 
      BorderWidth     =   5
      Index           =   10
      X1              =   7320
      X2              =   7320
      Y1              =   3000
      Y2              =   3300
   End
   Begin VB.Line W 
      BorderWidth     =   5
      Index           =   9
      X1              =   7320
      X2              =   9220
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line W 
      BorderWidth     =   5
      Index           =   8
      X1              =   9000
      X2              =   9000
      Y1              =   3300
      Y2              =   5000
   End
   Begin VB.Line W 
      BorderWidth     =   5
      Index           =   7
      X1              =   7320
      X2              =   9220
      Y1              =   3300
      Y2              =   3300
   End
   Begin VB.Line W 
      BorderWidth     =   5
      Index           =   6
      X1              =   7500
      X2              =   7500
      Y1              =   3300
      Y2              =   5000
   End
   Begin VB.Line W 
      BorderWidth     =   5
      Index           =   5
      X1              =   3200
      X2              =   3200
      Y1              =   3000
      Y2              =   3300
   End
   Begin VB.Line W 
      BorderWidth     =   5
      Index           =   4
      X1              =   1300
      X2              =   1300
      Y1              =   3000
      Y2              =   3300
   End
   Begin VB.Line W 
      BorderWidth     =   5
      Index           =   3
      X1              =   1300
      X2              =   3200
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line W 
      BorderWidth     =   5
      Index           =   2
      X1              =   3000
      X2              =   3000
      Y1              =   3300
      Y2              =   5000
   End
   Begin VB.Line W 
      BorderWidth     =   5
      Index           =   1
      X1              =   1300
      X2              =   3200
      Y1              =   3300
      Y2              =   3300
   End
   Begin VB.Line W 
      BorderWidth     =   5
      Index           =   0
      X1              =   1500
      X2              =   1500
      Y1              =   3300
      Y2              =   5000
   End
   Begin VB.Line Ground 
      BorderWidth     =   5
      X1              =   0
      X2              =   12000
      Y1              =   5000
      Y2              =   5000
   End
   Begin VB.Image Head 
      Height          =   1212
      Left            =   3360
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Line Arm 
      BorderWidth     =   2
      X1              =   840
      X2              =   1320
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Body 
      BorderWidth     =   6
      X1              =   1080
      X2              =   1080
      Y1              =   1560
      Y2              =   2040
   End
   Begin VB.Line RFeet 
      BorderWidth     =   3
      X1              =   1080
      X2              =   1320
      Y1              =   2040
      Y2              =   2400
   End
   Begin VB.Line LFeet 
      BorderWidth     =   3
      X1              =   1080
      X2              =   840
      Y1              =   2040
      Y2              =   2400
   End
End
Attribute VB_Name = "Egg99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Shaked As Boolean, Shake As Integer, Shake2 As Integer
Const ArmS As Integer = 20, FeetS As Integer = 20, ShakeInt As Integer = 30
Dim Nm As Boolean

Private Type Player
    X As Long
    y As Long
End Type

Dim Gao As Integer, Ai As Integer, Pang As Integer
Dim I As Player
Dim z As Boolean, y As Boolean, s As Boolean
Dim Grounded As Boolean, vSpeed As Integer
Const WalkSpeed As Integer = 40
Const Gravity As Integer = 15
Const Jump As Integer = 280

Private Sub Form_GotFocus()
Focuser.SetFocus
End Sub

Private Sub Focuser_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37     '左
        z = True
    Case 38     '上
        s = True
    Case 39    '右
        y = True
End Select
End Sub

Private Sub Focuser_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37     '左
        z = False
    Case 38     '上
        s = False
    Case 39    '右
        y = False
End Select
End Sub
Private Sub Form_Load()
Nm = False
Head.Picture = LoadPicture(App.Path & "\R\face.jpg")
Head.Top = Body.Y1 - Head.Height
Head.Left = Body.X1 - Head.Width / 2
I.y = 3600
I.X = 800
Pang = Head.Width / 2
Gao = Head.Height
Ai = 480
Shaker.Interval = ShakeInt
Main.Interval = 20
Show
Focuser.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not Nm Then
    End
End If
End Sub

Private Sub Main_Timer()
Select Case I.X
    Case 0 To 1500 - Pang '左管左侧
        I.X = I.X + (z - y) * WalkSpeed '左右移动
        '
        If I.X <= 0 Then
            I.X = 1
        End If
        If I.X >= 1500 - Pang Then
            If I.y + Ai < 3000 Then
                '成功跳过
            Else    '被挡住
                I.X = 1500 - Pang
            End If
        End If
        '
        vSpeed = vSpeed + Gravity
        I.y = I.y + vSpeed
        '
        If I.y + Ai >= 5000 Then '着地
            Grounded = True
            vSpeed = 0
            I.y = 5000 - Ai
        Else
            Grounded = False
        End If
        '
        If Grounded And s Then  '跳跃
            vSpeed = -Jump
        End If
        '
    Case 1501 - Pang To 3000 + Pang '左管上
        I.X = I.X + (z - y) * WalkSpeed '左右移动
        If I.X <= 1501 - Pang Or I.X >= 3000 + Pang Then '左侧下落
            GoTo TiaoGuo
        End If
        '
        vSpeed = vSpeed + Gravity
        I.y = I.y + vSpeed
        '
        If I.y + Ai >= 3000 Then '着地
            Grounded = True
            vSpeed = 0
            I.y = 3000 - Ai
        Else
            Grounded = False
        End If
        '
        If Grounded And s Then  '跳跃
            vSpeed = -Jump
        End If
        '
    Case 3001 + Pang To 7500 - Pang '中间地上
        I.X = I.X + (z - y) * WalkSpeed '左右移动
        '
        If I.X <= 3001 + Pang Then
            If I.y + Ai < 3000 Then
                '成功跳过
            Else
                I.X = 3002 + Pang
            End If
        End If
        If I.X >= 7500 - Pang Then
            If I.y + Ai < 3000 Then
                '成功跳过
            Else    '被挡住
                I.X = 7499 - Pang
            End If
        End If
        '
        vSpeed = vSpeed + Gravity
        I.y = I.y + vSpeed
        '
        If I.y + Ai >= 5000 Then '着地
            Grounded = True
            vSpeed = 0
            I.y = 5000 - Ai
        Else
            Grounded = False
        End If
        '
        If Grounded And s Then  '跳跃
            vSpeed = -Jump
        End If
        '
    Case 7501 = Pang To 9000 + Pang '右管上
        I.X = I.X + (z - y) * WalkSpeed '左右移动
        If I.X <= 7501 - Pang Or I.X >= 9000 + Pang Then '左侧下落
            GoTo TiaoGuo
        End If
        '
        vSpeed = vSpeed + Gravity
        I.y = I.y + vSpeed
        '
        If I.y + Ai >= 3000 Then '着地
            Grounded = True
            vSpeed = 0
            I.y = 3000 - Ai
        Else
            Grounded = False
        End If
        '
        If Grounded And s Then  '跳跃
            vSpeed = -Jump
        End If
        '
    Case Is >= 9001 + Pang  '右管右
        I.X = I.X + (z - y) * WalkSpeed '左右移动
        '
        If I.X <= 9001 + Pang Then
            If I.y + Ai <= 3000 Then
                '成功跳过
            Else
                I.X = 9002 + Pang
            End If
        End If
        If I.X >= Me.Width Then
            I.X = Me.Width - 1
        End If
        '
        vSpeed = vSpeed + Gravity
        I.y = I.y + vSpeed
        '
        If I.y + Ai >= 5000 Then '着地
            Grounded = True
            vSpeed = 0
            I.y = 5000 - Ai
        Else
            Grounded = False
        End If
        '
        If Grounded And s Then  '跳跃
            vSpeed = -Jump
        End If
        '
End Select

''''''''
TiaoGuo:

If I.X + Pang >= Q.Left And I.X - Pang <= Q.Left + Q.Width And I.y + Ai >= Q.Top And I.y - Gao <= Q.Top + Q.Height Then
    '碰到|?|
    If I.X + Pang - Q.Left <= 100 Then
        '撞左墙
        I.X = Q.Left - Pang - 5
        GoTo TiaoGuo2
    End If
    If Q.Left + Q.Width - I.X + Pang <= 100 Then
        '撞右墙
        I.X = Q.Left + Q.Width + Pang + 5
        GoTo TiaoGuo2
    End If
    If I.y + Ai - Q.Top <= 100 Then
        '踩到？
        vSpeed = 0
        I.y = Q.Top - Ai
        Grounded = True
    End If
    If Q.Top + Q.Height - I.y + Gao <= 100 Then
        '顶到？
        vSpeed = 10 '弹性
        I.y = Q.Top + Q.Height + Gao
    End If
End If
'''''''''''''
TiaoGuo2:

LFeet.Y2 = I.y + Ai
RFeet.Y2 = LFeet.Y2
LFeet.Y1 = I.y + 120
RFeet.Y1 = LFeet.Y1
LFeet.X1 = I.X
RFeet.X1 = LFeet.X1
With Body
    .X1 = I.X
    .X2 = I.X
    .Y1 = I.y - 120
    .Y2 = I.y + 120
End With
Arm.X1 = I.X - 240
Arm.X2 = I.X + 240
Head.Left = I.X - Pang
'
If I.X > Win.Left And I.X < Win.Left + Win.Width And I.y > Win.Top And I.y < Win.Top + Win.Height Then
    '碰到了菱形
    MsgBox "恭喜你通过了第2关。", vbInformation + vbOKOnly, "赢了！"
    Nm = True
    Load Egg3
    Unload Me
End If
End Sub

Private Sub Shaker_Timer()
Shake = (Shake + 1) Mod 20
Shake2 = (Shake2 + 1) Mod 60
Select Case Shake
    Case 0 To 5
        Arm.Y1 = I.y + ArmS * Shake
        Arm.Y2 = I.y - ArmS * Shake
    Case 6 To 15
        Arm.Y1 = I.y + 10 * ArmS - Shake * ArmS
        Arm.Y2 = I.y - 10 * ArmS + Shake * ArmS
    Case 16 To 19
        Arm.Y1 = I.y - 20 * ArmS + Shake * ArmS
        Arm.Y2 = I.y + 20 * ArmS - Shake * ArmS
End Select
If Shake2 < 30 Then
    RFeet.X2 = I.X + 15 * FeetS - Shake2 * FeetS
    LFeet.X2 = I.X - 15 * FeetS + Shake2 * FeetS
Else
    RFeet.X2 = I.X - 45 * FeetS + Shake2 * FeetS
    LFeet.X2 = I.X + 45 * FeetS - Shake2 * FeetS
End If
    Head.Top = I.y - (Body.Y2 - Body.Y1) / 2 - Head.Height + Abs(Shake2 - 30) * 5
End Sub

Private Sub Win_GotFocus()
Show
Focuser.SetFocus
End Sub

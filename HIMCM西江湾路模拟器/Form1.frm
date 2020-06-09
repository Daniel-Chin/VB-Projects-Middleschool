VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   7440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9876
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9876
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Main 
      Left            =   5280
      Top             =   2880
   End
   Begin VB.Shape QQ 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   200
      Left            =   8160
      Top             =   5040
      Width           =   200
   End
   Begin VB.Shape Gate 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      FillColor       =   &H0000FF00&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1452
      Left            =   3240
      Top             =   4920
      Width           =   4212
   End
   Begin VB.Line LR 
      BorderWidth     =   2
      X1              =   0
      X2              =   360
      Y1              =   0
      Y2              =   2160
   End
   Begin VB.Line LL 
      BorderWidth     =   2
      X1              =   1080
      X2              =   1440
      Y1              =   1200
      Y2              =   3360
   End
   Begin VB.Shape Da 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      FillStyle       =   6  'Cross
      Height          =   2052
      Index           =   0
      Left            =   6840
      Top             =   1680
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.Label Che 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1572
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1812
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'模拟实验校门口的交通情况。
'程序注释中，“'d”表示可调控参数。
'长度单位是 米。

Option Explicit

Dim Time As Long

Private Type Consts '常量
    Return As String '回车字符
    CarSize As Single '轿车长度
    CarWidth As Single '轿车宽度
    BusSize As Single '巴士长度
    BusWidth As Single '巴士宽度
    DamnTime As Integer '路障存在时间
    SpeedLimit As Single '限速
    MapSize As Single   '程序模拟马路范围(米)
    GateStart As Single '学校大门A端
    GateEnd As Single '学校大门B端
    ParkTime As Byte '送孩子时家长停在路边的时间
    RoadWidth As Single '路宽
    JiaSuDu As Single '车辆最高加速度
End Type

Dim Con As Consts
 
Private Type Cars
    Target As Byte  '目的地
    Road As Byte    '所在道路
    Position As Single  '所在位置
    Selfishness As Byte '自私度
    Anxiousness As Byte '心情急切程度
    Exist As Boolean    '存在性
    Size As Single  '多长
    Width As Single
    IsParking As Boolean '在停车吗
    ParkingCountDown As Integer '停车倒计时
    SafeDistance As Single '安全车距
    Speed As Single '车速
End Type

Private Type damns
    Road As Byte
    Position As Single
    TimeCountDown As Integer    '消失倒计时
    Exist As Boolean
End Type

Private Type Point '点
    X As Long
    Y As Long
End Type

Private Type KeShiHua '可视化参数
    Q As Point '地图原点位置
    Zoom As Single '放大比率
    MoveSpeed As Long '镜头移动速度
End Type

Dim Vis As KeShiHua

Dim k As Byte   '循环变量
Dim Car(0 To 254) As Cars, Damn(0 To 254) As damns

Private Sub Form_KeyPress(KeyAscii As Integer)
With Vis
    Select Case KeyAscii
        Case 113 'q
            If Main.Interval <> 0 Then
                Main.Interval = Main.Interval - 1
            End If
            Cls
            Print "Interval = " & Main.Interval
        Case 101 'e
            Main.Interval = Main.Interval + 1
            Cls
            Print "Interval = " & Main.Interval
        Case 119 'w
            .Q.Y = .Q.Y + .MoveSpeed
        Case 115 's
            .Q.Y = .Q.Y - .MoveSpeed
        Case 97 'a
            .Q.X = .Q.X - .MoveSpeed
        Case 100 'd
            .Q.X = .Q.X + .MoveSpeed
        Case 111 'o
            .Zoom = .Zoom + 0.6
        Case 108 'l
            .Zoom = .Zoom - 0.6
        Case Else
            Exit Sub
    End Select
End With
Renew
End Sub

Private Sub Form_Load()
MsgBox "WASD控制镜头，OL调节放大比率，QE控制时间流逝速度"
Time = 0
Randomize
'定义常量
With Con
    .Return = Chr(13) & Chr(10)
    .BusSize = 12
    .BusWidth = 2.5
    .CarWidth = 1.8
    .CarSize = 5
    .DamnTime = 10 'd
    .SpeedLimit = 5 'd
    .MapSize = 120 'd
    .GateEnd = 80 'd
    .GateStart = 70 'd
    .ParkTime = 26 'd
    .RoadWidth = 3 'd
    .JiaSuDu = 0.1 'd
End With
With Vis
    .Q.X = Me.Width * 4 / 5
    .Q.Y = Me.Height * 4 / 5
    .Zoom = Me.Height / Con.MapSize
    .MoveSpeed = 100
End With

RenewRoad
Main.Interval = 18 'd
End Sub

Private Sub Main_Timer()    '时间的进行
Time = Time + 1
'If Time = 177 Then Stop
'一次timer相当于1/3秒
'Debug.Print "s"
'生成车辆
If Rnd(1) <= 0.26 Then      'd控制生成车的频率
    Dim NewCarID As Byte '生成的车辆的ID
    For k = 1 To 254    '看看哪个车辆ID是空着的
        If Car(k).Exist = False Then
            NewCarID = k '找到了！那么新生成车的ID就用你了
            Exit For
        End If
    Next k
    If NewCarID = 0 Then '万一没找到捏？
        MsgBox "车辆超出了上限！", vbCritical + vbOKOnly, "警告"
        Stop
        Main.Interval = 0
        Exit Sub
    End If
    With Car(NewCarID) '生成的车是什么样的呢？
        .Exist = True '首先，它存在。
        .Anxiousness = 0 '出场的时候，大家都心平气和
        .Selfishness = Int(Rnd() * 256) '随机的自私度
        .SafeDistance = 0.5 '安全车距
        If Rnd() <= 0.5 Then 'd这辆车有一定的几率...
            '是来送孩子的！
            .Target = 1
            .Size = Con.CarSize '一定是小轿车
            .Width = Con.CarWidth
            .Road = Int(Rnd * 3) + 1 '在哪条路生成？
            '决定一下它的生成位置。
            '首先找到这条路上最底下的车(或damn)的位置
            Dim Bond As Single
            Bond = 5 '定义暂存
            For k = 1 To 254
                If Car(k).Exist And Car(k).Road = .Road Then
                    '如果被选定的车存在且在这条路上
                    If Car(k).Position < Bond Then
                        '与暂存比较
                        Bond = Car(k).Position
                    End If
                End If
                'damn再来一次
                If Damn(k).Exist And Damn(k).Road = .Road Then
                      If Damn(k).Position < Bond Then
                        Bond = Damn(k).Position
                    End If
                End If
            Next k
            .Position = Bond '这样他就到最底下了
            .Position = .Position - 2 * Con.BusSize - 0.5 'd留出距离——他不能和其他车重合啊！
        Else
            '是通过西江湾路的！
            .Target = 0
            .Road = Int(Rnd * 3) + 1 '在哪条路生成？
            If .Selfishness <= 50 Then
                .Road = 3  '良心好的人会在外侧登场
            End If
            '决定一下它的生成位置。
            '首先找到这条路上最底下的车的位置
            Bond = 5 '定义暂存
            For k = 1 To 254
                If Car(k).Exist And Car(k).Road = .Road Then
                    '如果被选定的车存在且在这条路上
                    If Car(k).Position < Bond Then
                        '与暂存比较
                        Bond = Car(k).Position
                    End If
                End If
                If Damn(k).Exist And Damn(k).Road = .Road Then
                      If Damn(k).Position < Bond Then
                        Bond = Damn(k).Position
                    End If
                End If
            Next k
            .Position = Bond  '这样他就到最底下了
            .Position = .Position - 2 * Con.BusSize - 0.5 'd留出距离——他不能和其他车重合啊！
            '接下来决定它的长度
            If Rnd() <= 0.2 Then 'd它有一定几率...
                '是大巴！
                .Size = Con.BusSize
                .Width = Con.BusWidth
            Else
                '是轿车！
                .Size = Con.CarSize
                .Width = Con.CarWidth
            End If
        End If
        '这样，这辆车就存在于内存中了。
        '接下来要把它显示出来
        Load Che(NewCarID)
        With Che(NewCarID)
            .BackColor = RGB(Rnd() * 100, Rnd() * 100, Rnd() * 100)
            RenewChe NewCarID
            .Visible = True
        End With
    End With
End If

'消除抵达的车辆
For k = 1 To 254 '把所有车循环一遍
    With Car(k)
        If .Exist Then '选定的车辆如果存在
            If .Position >= Con.MapSize Then '如果车辆位置超出地图边界
                .Exist = False '灭了它
                Unload Che(k)
                'Debug.Print "del"; k
            End If
        End If
    End With
Next k

'接下来要把所有车循环一遍，看看他们的行为。
Dim CarID As Byte
For CarID = 1 To 254
    With Car(CarID)
'        If CarID = 11 Then Stop
        If .Exist Then    '它存在吗
            If .IsParking Then
                '在停车
                .ParkingCountDown = .ParkingCountDown - 1
                If .ParkingCountDown <= 0 Then
                    .IsParking = False
                    .Target = 0 '接着驶出
                End If
            Else
                '不在停车
                '首先计算安全车距
                .SafeDistance = V2S(.Speed) + 0.1
                '计算前面一辆车车尾位置
                Bond = Con.MapSize + Con.BusSize '定义暂存
                For k = 1 To 254
                    If Car(k).Exist And Car(k).Position > .Position And Car(k).Road = .Road Then
                        If Car(k).Position < Bond Then
                            Bond = Car(k).Position
                        End If
                    End If
                    If Damn(k).Exist And Damn(k).Position > .Position And Damn(k).Road = .Road Then
                        If Damn(k).Position < Bond Then
                            Bond = Damn(k).Position
                        End If
                    End If
                Next k
                Dim PiGu As Single '车尾位置
                PiGu = Bond
                '再计算两边变道的可行性
                Dim ToL As Boolean, ToR As Boolean 'L for Left,R for Right
                '先计算ToL
                If .Road = 3 Then
                    ToL = False '最左边无法再向左
                Else
                    '不是最左边
                    '检测重合
                    ToL = True
                    For k = 1 To 254
                        If Car(k).Exist And Car(k).Road = .Road + 1 Then
                            '被选中的车在左边的路上
                            If COL(CarID, k) Then
                                '撞了
                                ToL = False
                                Exit For
                            End If
                        End If
                        '再算damn
                        If Damn(k).Exist And Damn(k).Road = .Road + 1 Then
                            '被选中的damn在左边的路上
                            If COLD(CarID, k) Then
                                '撞了
                                ToL = False
                                Exit For
                            End If
                        End If
                    Next k
                End If
                '再计算ToR
                If .Road = 1 Then
                    ToR = False '最右边无法再向右
                Else
                    '不是最右边
                    '检测重合
                    ToR = True
                    For k = 1 To 254
                        If Car(k).Exist And Car(k).Road = .Road - 1 Then
                            '被选中的车在左边的路上
                            If COL(CarID, k) Then
                                '撞了
                                ToR = False
                                Exit For
                            End If
                        End If
                        '再算damn
                        If Damn(k).Exist And Damn(k).Road = .Road + 1 Then
                            '被选damn的车在左边的路上
                            If COLD(CarID, k) Then
                                '撞了
                                ToL = False
                                Exit For
                            End If
                        End If
                    Next k
                End If
                'PiGu,ToL,ToR都计算完毕。
                If .Target = 0 Then '他是干么的？
Ahead:              '通过西江湾路。
                    '车辆前行
                    Dim Wanted As Single '期望到达的位置
                    Wanted = .Position + .Speed
                    If Wanted + .Size / 2 + .SafeDistance >= PiGu Then
                        '堵车
                        .Speed = .Speed / 2
                        .Speed = Abs(.Speed)
                        .Position = .Position + .Speed
                        Dim Anger As Integer
                        Anger = Int((Wanted - .Position) * 6) + 1
                        If Anger + .Anxiousness <= 255 Then
                            .Anxiousness = .Anxiousness + Anger 'd达不到预期位置，郁闷啊
                        End If
                    Else
                        '畅通
                        .Position = Wanted
                        .Speed = .Speed + Con.JiaSuDu
                        If .Speed >= Con.SpeedLimit Then
                            '不能超速
                            .Speed = Con.SpeedLimit
                        End If
                        If .Anxiousness <> 0 Then
                            .Anxiousness = .Anxiousness - 1 '消气
                        End If
                    End If
                    '这车会不会变道呢？
                    If .Anxiousness / 330 + .Selfishness / 300 >= Rnd() And .Anxiousness <> 0 Then 'd论文模型一中的函数h出现了
                        '他决定要变道!
                        If ToL Or ToR Then
                            If Rnd() <= 0.5 Then '有几率
TR:                             '他要向右
                                If ToR Then
                                    LoadDamn .Road, .Position
                                    .Road = .Road - 1
                                    .Anxiousness = 0 '消气了
                                Else
                                    GoTo TL
                                End If
                            Else
TL:                             '他要向左
                                If ToL Then
                                    LoadDamn .Road, .Position
                                    .Road = .Road + 1
                                    .Anxiousness = 0 '消气了
                                Else
                                    GoTo TR
                                End If
                            End If
                            .Speed = 0
                        End If
                    End If
                Else
                    '接送孩子。
                    Select Case .Position
                        Case Is < Con.GateStart
                            '还没到校门
                            GoTo Ahead
                        Case Con.GateStart To Con.GateEnd
                            '在校门范围，可以试着靠过去停车了！
                            If ToR Then '可以向右变道吗？
                                LoadDamn .Road, .Position
                                .Road = .Road - 1
                            Else
                                If .Road = 1 Then '如果已经在最右边了
                                    .IsParking = True '停车下客
                                End If
                                GoTo Ahead
                            End If
                        Case Is > Con.GateEnd
                            '超过了校门范围，还没有把孩子放下来！
                            .IsParking = True '直接停车
                            .ParkingCountDown = Con.ParkTime
                    End Select
                End If
            End If
            'Debug.Print CarID; .Road; .Position
            RenewChe CarID
        End If
    End With
Next CarID
'所有车辆行为完毕。
'接下来看那些damn。
Dim DamnID As Byte
For DamnID = 1 To 254
    With Damn(DamnID)
        If .Exist Then
            .TimeCountDown = .TimeCountDown - 1
            If .TimeCountDown <= 0 Then
                DeloadDamn DamnID
            End If
        End If
    End With
Next DamnID
'damn计算完毕
'这1/3秒终于过去了。

'设计员注释：
'图形输出
End Sub

Private Function COL(ID1 As Byte, ID2 As Byte) As Boolean    '是否重合
COL = Car(ID1).Position - Car(ID1).Size / 2 <= Car(ID2).Position + Car(ID2).Size / 2 And Car(ID1).Position + Car(ID1).Size / 2 > Car(ID2).Position - Car(ID2).Size / 2
End Function

Private Function COLD(CarID As Byte, DamnID As Byte) As Boolean    '是否重合
COLD = Car(CarID).Position - Car(CarID).Size / 2 <= Damn(DamnID).Position + Con.CarSize / 2 And Car(CarID).Position + Car(CarID).Size / 2 > Damn(DamnID).Position - Con.CarSize / 2
End Function

Private Sub LoadDamn(Road As Byte, Position As Single)
'生成Damn。
Dim DamnID As Byte
For k = 1 To 254 '把所有damn循环一遍
    If Not Damn(k).Exist Then
        DamnID = k
        Exit For
    End If
Next k
'DamnID已决定
With Damn(DamnID)
    .Exist = True
    .Road = Road
    .Position = Position
    .TimeCountDown = Con.DamnTime
End With
Load Da(DamnID)
RenewDa DamnID
Da(DamnID).Visible = True
End Sub

Private Sub DeloadDamn(ID As Byte)
'卸载damn
Damn(ID).Exist = False
Unload Da(ID)
End Sub

Private Function V2S(Speed As Single) As Single
'根据车速计算安全车距。
V2S = Speed * 2 'd
End Function

Private Sub RenewChe(ID As Byte)
Debug.Assert Car(ID).Exist
With Che(ID)
    .Left = Vis.Q.X - Vis.Zoom * ((Car(ID).Road - 0.5) * Con.RoadWidth + Car(ID).Width / 2)
    .Width = Vis.Zoom * Car(ID).Width
    .Top = Vis.Q.Y - Vis.Zoom * (Car(ID).Size / 2 + Car(ID).Position)
    .Height = Vis.Zoom * Car(ID).Size
    .Caption = ID
End With
End Sub

Private Sub RenewDa(ID As Byte)
Debug.Assert Damn(ID).Exist
With Da(ID)
    .Left = Vis.Q.X - Vis.Zoom * ((Damn(ID).Road - 0.5) * Con.RoadWidth + Con.CarWidth / 2)
    .Width = Vis.Zoom * Con.CarWidth
    .Top = Vis.Q.Y - Vis.Zoom * (Con.CarSize / 2 + Damn(ID).Position)
    .Height = Vis.Zoom * Con.CarSize
End With
End Sub

Private Sub RenewRoad()
LL.X1 = Vis.Q.X - Vis.Zoom * (Con.RoadWidth) * 2
LL.X2 = LL.X1
LL.Y1 = 0
LL.Y2 = Me.Height
LR.X1 = Vis.Q.X - Vis.Zoom * (Con.RoadWidth)
LR.X2 = LR.X1
LR.Y1 = 0
LR.Y2 = Me.Height
With Gate
    .Left = 0
    .Width = Me.Width
    .Top = Vis.Q.Y - Vis.Zoom * (Con.GateEnd)
    .Height = Vis.Zoom * (Con.GateEnd - Con.GateStart)
End With
QQ.Top = Vis.Q.Y
QQ.Left = Vis.Q.X
End Sub

Private Sub Renew() '更新所有，调试用
For k = 1 To 254
    If Car(k).Exist Then
        RenewChe k
    End If
Next k
RenewRoad
End Sub

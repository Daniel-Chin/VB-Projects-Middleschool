VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   7440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9876
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9876
   StartUpPosition =   3  '����ȱʡ
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
'ģ��ʵ��У�ſڵĽ�ͨ�����
'����ע���У���'d����ʾ�ɵ��ز�����
'���ȵ�λ�� �ס�

Option Explicit

Dim Time As Long

Private Type Consts '����
    Return As String '�س��ַ�
    CarSize As Single '�γ�����
    CarWidth As Single '�γ����
    BusSize As Single '��ʿ����
    BusWidth As Single '��ʿ���
    DamnTime As Integer '·�ϴ���ʱ��
    SpeedLimit As Single '����
    MapSize As Single   '����ģ����·��Χ(��)
    GateStart As Single 'ѧУ����A��
    GateEnd As Single 'ѧУ����B��
    ParkTime As Byte '�ͺ���ʱ�ҳ�ͣ��·�ߵ�ʱ��
    RoadWidth As Single '·��
    JiaSuDu As Single '������߼��ٶ�
End Type

Dim Con As Consts
 
Private Type Cars
    Target As Byte  'Ŀ�ĵ�
    Road As Byte    '���ڵ�·
    Position As Single  '����λ��
    Selfishness As Byte '��˽��
    Anxiousness As Byte '���鼱�г̶�
    Exist As Boolean    '������
    Size As Single  '�೤
    Width As Single
    IsParking As Boolean '��ͣ����
    ParkingCountDown As Integer 'ͣ������ʱ
    SafeDistance As Single '��ȫ����
    Speed As Single '����
End Type

Private Type damns
    Road As Byte
    Position As Single
    TimeCountDown As Integer    '��ʧ����ʱ
    Exist As Boolean
End Type

Private Type Point '��
    X As Long
    Y As Long
End Type

Private Type KeShiHua '���ӻ�����
    Q As Point '��ͼԭ��λ��
    Zoom As Single '�Ŵ����
    MoveSpeed As Long '��ͷ�ƶ��ٶ�
End Type

Dim Vis As KeShiHua

Dim k As Byte   'ѭ������
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
MsgBox "WASD���ƾ�ͷ��OL���ڷŴ���ʣ�QE����ʱ�������ٶ�"
Time = 0
Randomize
'���峣��
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

Private Sub Main_Timer()    'ʱ��Ľ���
Time = Time + 1
'If Time = 177 Then Stop
'һ��timer�൱��1/3��
'Debug.Print "s"
'���ɳ���
If Rnd(1) <= 0.26 Then      'd�������ɳ���Ƶ��
    Dim NewCarID As Byte '���ɵĳ�����ID
    For k = 1 To 254    '�����ĸ�����ID�ǿ��ŵ�
        If Car(k).Exist = False Then
            NewCarID = k '�ҵ��ˣ���ô�����ɳ���ID��������
            Exit For
        End If
    Next k
    If NewCarID = 0 Then '��һû�ҵ���
        MsgBox "�������������ޣ�", vbCritical + vbOKOnly, "����"
        Stop
        Main.Interval = 0
        Exit Sub
    End If
    With Car(NewCarID) '���ɵĳ���ʲô�����أ�
        .Exist = True '���ȣ������ڡ�
        .Anxiousness = 0 '������ʱ�򣬴�Ҷ���ƽ����
        .Selfishness = Int(Rnd() * 256) '�������˽��
        .SafeDistance = 0.5 '��ȫ����
        If Rnd() <= 0.5 Then 'd��������һ���ļ���...
            '�����ͺ��ӵģ�
            .Target = 1
            .Size = Con.CarSize 'һ����С�γ�
            .Width = Con.CarWidth
            .Road = Int(Rnd * 3) + 1 '������·���ɣ�
            '����һ����������λ�á�
            '�����ҵ�����·������µĳ�(��damn)��λ��
            Dim Bond As Single
            Bond = 5 '�����ݴ�
            For k = 1 To 254
                If Car(k).Exist And Car(k).Road = .Road Then
                    '�����ѡ���ĳ�������������·��
                    If Car(k).Position < Bond Then
                        '���ݴ�Ƚ�
                        Bond = Car(k).Position
                    End If
                End If
                'damn����һ��
                If Damn(k).Exist And Damn(k).Road = .Road Then
                      If Damn(k).Position < Bond Then
                        Bond = Damn(k).Position
                    End If
                End If
            Next k
            .Position = Bond '�������͵��������
            .Position = .Position - 2 * Con.BusSize - 0.5 'd�������롪�������ܺ��������غϰ���
        Else
            '��ͨ��������·�ģ�
            .Target = 0
            .Road = Int(Rnd * 3) + 1 '������·���ɣ�
            If .Selfishness <= 50 Then
                .Road = 3  '���ĺõ��˻������ǳ�
            End If
            '����һ����������λ�á�
            '�����ҵ�����·������µĳ���λ��
            Bond = 5 '�����ݴ�
            For k = 1 To 254
                If Car(k).Exist And Car(k).Road = .Road Then
                    '�����ѡ���ĳ�������������·��
                    If Car(k).Position < Bond Then
                        '���ݴ�Ƚ�
                        Bond = Car(k).Position
                    End If
                End If
                If Damn(k).Exist And Damn(k).Road = .Road Then
                      If Damn(k).Position < Bond Then
                        Bond = Damn(k).Position
                    End If
                End If
            Next k
            .Position = Bond  '�������͵��������
            .Position = .Position - 2 * Con.BusSize - 0.5 'd�������롪�������ܺ��������غϰ���
            '�������������ĳ���
            If Rnd() <= 0.2 Then 'd����һ������...
                '�Ǵ�ͣ�
                .Size = Con.BusSize
                .Width = Con.BusWidth
            Else
                '�ǽγ���
                .Size = Con.CarSize
                .Width = Con.CarWidth
            End If
        End If
        '�������������ʹ������ڴ����ˡ�
        '������Ҫ������ʾ����
        Load Che(NewCarID)
        With Che(NewCarID)
            .BackColor = RGB(Rnd() * 100, Rnd() * 100, Rnd() * 100)
            RenewChe NewCarID
            .Visible = True
        End With
    End With
End If

'�����ִ�ĳ���
For k = 1 To 254 '�����г�ѭ��һ��
    With Car(k)
        If .Exist Then 'ѡ���ĳ����������
            If .Position >= Con.MapSize Then '�������λ�ó�����ͼ�߽�
                .Exist = False '������
                Unload Che(k)
                'Debug.Print "del"; k
            End If
        End If
    End With
Next k

'������Ҫ�����г�ѭ��һ�飬�������ǵ���Ϊ��
Dim CarID As Byte
For CarID = 1 To 254
    With Car(CarID)
'        If CarID = 11 Then Stop
        If .Exist Then    '��������
            If .IsParking Then
                '��ͣ��
                .ParkingCountDown = .ParkingCountDown - 1
                If .ParkingCountDown <= 0 Then
                    .IsParking = False
                    .Target = 0 '����ʻ��
                End If
            Else
                '����ͣ��
                '���ȼ��㰲ȫ����
                .SafeDistance = V2S(.Speed) + 0.1
                '����ǰ��һ������βλ��
                Bond = Con.MapSize + Con.BusSize '�����ݴ�
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
                Dim PiGu As Single '��βλ��
                PiGu = Bond
                '�ټ������߱���Ŀ�����
                Dim ToL As Boolean, ToR As Boolean 'L for Left,R for Right
                '�ȼ���ToL
                If .Road = 3 Then
                    ToL = False '������޷�������
                Else
                    '���������
                    '����غ�
                    ToL = True
                    For k = 1 To 254
                        If Car(k).Exist And Car(k).Road = .Road + 1 Then
                            '��ѡ�еĳ�����ߵ�·��
                            If COL(CarID, k) Then
                                'ײ��
                                ToL = False
                                Exit For
                            End If
                        End If
                        '����damn
                        If Damn(k).Exist And Damn(k).Road = .Road + 1 Then
                            '��ѡ�е�damn����ߵ�·��
                            If COLD(CarID, k) Then
                                'ײ��
                                ToL = False
                                Exit For
                            End If
                        End If
                    Next k
                End If
                '�ټ���ToR
                If .Road = 1 Then
                    ToR = False '���ұ��޷�������
                Else
                    '�������ұ�
                    '����غ�
                    ToR = True
                    For k = 1 To 254
                        If Car(k).Exist And Car(k).Road = .Road - 1 Then
                            '��ѡ�еĳ�����ߵ�·��
                            If COL(CarID, k) Then
                                'ײ��
                                ToR = False
                                Exit For
                            End If
                        End If
                        '����damn
                        If Damn(k).Exist And Damn(k).Road = .Road + 1 Then
                            '��ѡdamn�ĳ�����ߵ�·��
                            If COLD(CarID, k) Then
                                'ײ��
                                ToL = False
                                Exit For
                            End If
                        End If
                    Next k
                End If
                'PiGu,ToL,ToR��������ϡ�
                If .Target = 0 Then '���Ǹ�ô�ģ�
Ahead:              'ͨ��������·��
                    '����ǰ��
                    Dim Wanted As Single '���������λ��
                    Wanted = .Position + .Speed
                    If Wanted + .Size / 2 + .SafeDistance >= PiGu Then
                        '�³�
                        .Speed = .Speed / 2
                        .Speed = Abs(.Speed)
                        .Position = .Position + .Speed
                        Dim Anger As Integer
                        Anger = Int((Wanted - .Position) * 6) + 1
                        If Anger + .Anxiousness <= 255 Then
                            .Anxiousness = .Anxiousness + Anger 'd�ﲻ��Ԥ��λ�ã����ư�
                        End If
                    Else
                        '��ͨ
                        .Position = Wanted
                        .Speed = .Speed + Con.JiaSuDu
                        If .Speed >= Con.SpeedLimit Then
                            '���ܳ���
                            .Speed = Con.SpeedLimit
                        End If
                        If .Anxiousness <> 0 Then
                            .Anxiousness = .Anxiousness - 1 '����
                        End If
                    End If
                    '�⳵�᲻�����أ�
                    If .Anxiousness / 330 + .Selfishness / 300 >= Rnd() And .Anxiousness <> 0 Then 'd����ģ��һ�еĺ���h������
                        '������Ҫ���!
                        If ToL Or ToR Then
                            If Rnd() <= 0.5 Then '�м���
TR:                             '��Ҫ����
                                If ToR Then
                                    LoadDamn .Road, .Position
                                    .Road = .Road - 1
                                    .Anxiousness = 0 '������
                                Else
                                    GoTo TL
                                End If
                            Else
TL:                             '��Ҫ����
                                If ToL Then
                                    LoadDamn .Road, .Position
                                    .Road = .Road + 1
                                    .Anxiousness = 0 '������
                                Else
                                    GoTo TR
                                End If
                            End If
                            .Speed = 0
                        End If
                    End If
                Else
                    '���ͺ��ӡ�
                    Select Case .Position
                        Case Is < Con.GateStart
                            '��û��У��
                            GoTo Ahead
                        Case Con.GateStart To Con.GateEnd
                            '��У�ŷ�Χ���������ſ���ȥͣ���ˣ�
                            If ToR Then '�������ұ����
                                LoadDamn .Road, .Position
                                .Road = .Road - 1
                            Else
                                If .Road = 1 Then '����Ѿ������ұ���
                                    .IsParking = True 'ͣ���¿�
                                End If
                                GoTo Ahead
                            End If
                        Case Is > Con.GateEnd
                            '������У�ŷ�Χ����û�аѺ��ӷ�������
                            .IsParking = True 'ֱ��ͣ��
                            .ParkingCountDown = Con.ParkTime
                    End Select
                End If
            End If
            'Debug.Print CarID; .Road; .Position
            RenewChe CarID
        End If
    End With
Next CarID
'���г�����Ϊ��ϡ�
'����������Щdamn��
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
'damn�������
'��1/3�����ڹ�ȥ�ˡ�

'���Աע�ͣ�
'ͼ�����
End Sub

Private Function COL(ID1 As Byte, ID2 As Byte) As Boolean    '�Ƿ��غ�
COL = Car(ID1).Position - Car(ID1).Size / 2 <= Car(ID2).Position + Car(ID2).Size / 2 And Car(ID1).Position + Car(ID1).Size / 2 > Car(ID2).Position - Car(ID2).Size / 2
End Function

Private Function COLD(CarID As Byte, DamnID As Byte) As Boolean    '�Ƿ��غ�
COLD = Car(CarID).Position - Car(CarID).Size / 2 <= Damn(DamnID).Position + Con.CarSize / 2 And Car(CarID).Position + Car(CarID).Size / 2 > Damn(DamnID).Position - Con.CarSize / 2
End Function

Private Sub LoadDamn(Road As Byte, Position As Single)
'����Damn��
Dim DamnID As Byte
For k = 1 To 254 '������damnѭ��һ��
    If Not Damn(k).Exist Then
        DamnID = k
        Exit For
    End If
Next k
'DamnID�Ѿ���
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
'ж��damn
Damn(ID).Exist = False
Unload Da(ID)
End Sub

Private Function V2S(Speed As Single) As Single
'���ݳ��ټ��㰲ȫ���ࡣ
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

Private Sub Renew() '�������У�������
For k = 1 To 254
    If Car(k).Exist Then
        RenewChe k
    End If
Next k
RenewRoad
End Sub

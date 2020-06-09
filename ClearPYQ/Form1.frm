VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   3744
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicBox 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   372
      Left            =   2640
      ScaleHeight     =   372
      ScaleWidth      =   972
      TabIndex        =   1
      Top             =   240
      Width           =   972
   End
   Begin VB.Timer Debugger 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   3240
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Label1"
      Height          =   132
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   132
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const CarefullyPick As Boolean = 1
Const Debuging As Boolean = 0
Const Speed As Integer = 80
Const DoBless As Boolean = 1
Const WaterPace As Integer = 240

Dim DBG As Integer

Dim HasShot As Boolean
Dim ShotPath As String
Const FadeDark As Long = 800
Const SluffBlue As Long = 8013331
Const SluffRed As Long = 3026585
Const ICenterX As Long = 1500
Const ICenterY As Long = 744
Const VDownstairR As Long = 480
Const Pi As Double = 3.14159265358979
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Dim OTS As Boolean
Dim BlessNum As Integer
Dim Slept As Long
Dim Hunger As Integer
Dim Wand As Single

Private Sub Snd(ByVal f As Long, Optional Time As Long = 100)
Beep f, Time
DoEvents
End Sub

Private Sub Music(Note As Integer, Optional Time As Long = 100)
Snd 440 * 2 ^ ((Note + 2) / 12), Time
End Sub

Private Sub Debugger_Timer()
Dim i As Double
If 1 Then
    Debugger.Interval = 3000
    Dim VD
    VD = VDownstair
    MsgBox VD
End If
If 0 Then
    Debugger.Interval = 30
    i = 0.442656606007002 + 0.01 * DBG
    Label1.Move ICenterX + Cos(i) * VDownstairR, ICenterY + Sin(i) * VDownstairR
    DBG = DBG + 1
End If
If 0 Then
    Debugger.Interval = 30
    i = Pi * 1.5
    Label1.Move ICenterX + Cos(i) * VDownstairR, ICenterY + Sin(i) * VDownstairR
    DBG = DBG + 1
End If
End Sub

Private Sub Form_Click()
DelayShot
End Sub

Private Sub Form_Load()
ShotPath = App.Path & "\1.bmp"
'Tag = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '-1 Do it,-2 Reset
'Show
Open "D:\vbbreak" For Binary As #1
Put #1, 1, 0
If Debuging Then
    Exit Sub
End If
Music 1
Snap 1000
Music 1
Snap 1000
Music 1
Snap 1000
Music 13
'WhatToDo?

Super
'WaterLoop
'StayHunt
'MakeHoney
'QuickUse
'GoDown
'Hunt
'Hatch
'TrainPet
'HuntressHunt
'Pitch
End Sub

Private Sub k(Txt As String, Optional SnapTime As Long = Speed, Optional DoCareBusy As Boolean = True)
Dim i As Integer
If Not Left(Txt, 1) = "{" And Len(Txt) >= 2 Then
    For i = 1 To Len(Txt)
        k Mid(Txt, i, 1), SnapTime, DoCareBusy
    Next i
Else
    Brake
    Dim myKey As Object
    Set myKey = CreateObject("WScript.Shell")
    myKey.SendKeys Txt
    Hunger = Hunger + 1
    Wand = Wand + 0.01
    If Hunger > 1000 Then Hunger = 1000
    Snap SnapTime / 2
    DoEvents
    If DoCareBusy Then
        Do While IsBusy
            Snap Speed
        Loop
    End If
    Snap SnapTime / 2
    Slept = 0
End If
End Sub

Private Sub Dian(Optional Act As Integer = 0)
Brake
Select Case Act
    Case 0
        mouse_event 6, 0, 0, 0, 0
    Case 1
        mouse_event 2, 0, 0, 0, 0
    Case -1
        mouse_event 4, 0, 0, 0, 0
End Select
End Sub

Private Sub Zhi(X As Long, Y As Long)
SetCursorPos X / 12, Y / 12
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If OTS Then MsgBox X & " " & Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #1
End Sub

Private Sub Pick()
    BlessNum = 0
    k "1"
    Snap 300
    StayPick
    BlessIfCan
    k "e"
    StayPick
    BlessIfCan
    k "a"
    StayPick
    BlessIfCan
    k "a"
    StayPick
    BlessIfCan
    k "s"
    StayPick
    BlessIfCan
    k "s"
    StayPick
    BlessIfCan
    k "d"
    StayPick
    BlessIfCan
    k "d"
    StayPick
    BlessIfCan
    k "w"
    StayPick
    k "a"
End Sub

Private Sub Water()
'    TrainCloak
Dim i As Integer
    Dim NoWater As Boolean
    'Do Until NoWater
    For i = 1 To 66
        k "{enter}", 100
        k "3", 100
        If CanWater() Then
            k "4"
            Snap 400
            '''''''''''''''''''
            k "956785678", WaterPace, False
            
'            k "567q", waterpace, False
            
            k "5678", WaterPace, False
            k "5678", WaterPace, False
            k "56e", WaterPace, False

            'Music -11
        Else
            'NoWater = True
            k "{enter}"
            k "{enter}"
            Exit For
        End If
    'Loop
    Next i
    'Music 1
End Sub

Private Sub Shot()
If OTS Then Exit Sub
    If Not HasShot Then
        DoEvents
        Dim w As Long, h As Long
        Dim hwndSrc As Long
        Dim hSrcDC As Long
        
        hwndSrc = GetDesktopWindow()
        hSrcDC = GetWindowDC(hwndSrc)
        w = Screen.Width \ Screen.TwipsPerPixelX
        h = Screen.Height \ Screen.TwipsPerPixelY
        
        Call BitBlt(hdc, 0, 0, w, h, hSrcDC, 0, 0, SRCCOPY)
        Call ReleaseDC(hwndSrc, hSrcDC)
        DoEvents
        HasShot = True
    End If
End Sub

Private Function FailSafe() As Boolean
Shot
Dim i As Long
Dim Cl As Long
For i = 42 To 23 Step -1
    Cl = Point(i * 100, 204)
    Cl = (Cl And &HFF&) - (Cl And &HFF00&) / &H100&
    If Cl >= 130 Then
        Exit For
    End If
Next i
If (i - 23) / (42 - 23) <= 1 / 3 Then
    FailSafe = True
End If
End Function

Private Sub Warn()
    Music 13, 100
    Music 13, 100
    Music 13, 100
    Music 13, 500
    Music 13, 500
    Stop
End Sub

Private Function CanWater() As Boolean
Shot
If Point(5844, 6708) <> SluffRed Then
    CanWater = True
End If
End Function

Private Function CanBless() As Boolean
Shot
If Point(9132, 6400) <> 3751998 Then
    CanBless = True
End If
End Function

Private Function SluffNotBlack() As Boolean
Shot
If Point(15492, 7956) <> 0 Then
    SluffNotBlack = True
End If
End Function

Private Sub StayPick()
If CarefullyPick Then
    Do While WhereBlueSluff
        k " "
    Loop
Else
    Do While SluffNotBlack()
        k " "
    Loop
End If
End Sub

Private Sub BlessIfCan()
'If BlessNum <= 1 Then
If DoBless Then
    Do
        k "{enter}"
        k "3"
        If CanBless() Then
            k "b"
            Music 13
            BlessNum = BlessNum + 1
        Else
            k "{enter}"
            k "{enter}"
            Exit Do
        End If
    Loop
End If
End Sub

Private Sub DelayShot()
If Debuging Then
    If OTS = False Then
        Snap 1000
        Snap 1000
        Snap 1000
        Shot
        OTS = True
    End If
End If
End Sub

Private Sub TrainCloak()
Dim i As Integer
k "2"
Snap 400
For i = 1 To 20
    k "f"
Next i
End Sub

Private Sub Timer1_Timer()
If FailSafe = True Then Music 1: Snap 1000
End Sub

Private Sub WaterLoop()
Do
    Water
    Pick
Loop
End Sub

Private Sub StayHunt()
Dim WRS As Integer, WBS As Integer
Do
    Do Until EnemyInSight(False, True)
        'Debug.Assert WhereRedSluff = 0
        'Debug.Assert WhereBlueSluff = 0
        Hunger = 0 '''''绝食
        If Hunger >= 400 Then
            k "{enter}"
            k "7"
            k "h"
            Hunger = 0
            Snap 700
        Else
            If Wand > 6 Then
                k "2"
                Snap 300
                k "d"
                Wand = Wand - 1
                Snap 300
            Else
                k "v", 200
                Wand = 0 'NoWand
            End If
        End If
    Loop
    'Music 1
    'Debug.Assert WRS <> 0
    Do Until EnemyInSight = False
        If 0 Then
            If WRS = 1 Then
                k "m"
            Else
                k "j"
            End If
        Else
            k "22"
        End If
        Snap 300
    Loop
    If 0 Then
            k "c"
            WBS = WhereBlueSluff
            Do Until WBS = 0
                k " "
                Snap 300
                WBS = WhereBlueSluff
            Loop
            k "q"
            Music 13
            Snap 300
    End If
    If 0 Then
        k "2"
        k "s"
    End If
Loop
End Sub

Private Function WhereRedSluff(Optional Arbitrary As Boolean = False) As Integer
Shot
If Point(15492, 7956) = SluffRed Then
    WhereRedSluff = 1
End If
If Point(15492, 6900) = SluffRed Then
    WhereRedSluff = 2
End If
If WhereRedSluff = 0 Then
    If Arbitrary Then
        If Point(15492, 7956) <> 0 Then
            WhereRedSluff = 1
        End If
        If Point(15492, 6900) <> 0 Then
            WhereRedSluff = 2
        End If
    Else
        Snap FadeDark - Slept
        Shot
        If Point(15492, 7956) = SluffRed Then
            WhereRedSluff = 1
        End If
        If Point(15492, 6900) = SluffRed Then
            WhereRedSluff = 2
        End If
    End If
End If
End Function

Private Function WhereBlueSluff() As Integer
Shot
If Point(15492, 7956) = SluffBlue Then
    WhereBlueSluff = 1
End If
If Point(15492, 6900) = SluffBlue Then
    WhereBlueSluff = 2
End If
If WhereBlueSluff = 0 Then
    Snap FadeDark - Slept
    Shot
    If Point(15492, 7956) = SluffBlue Then
        WhereBlueSluff = 1
    End If
    If Point(15492, 6900) = SluffBlue Then
        WhereBlueSluff = 2
    End If
End If
End Function

Private Function EnemyInSight(Optional ArbitraryNo As Boolean = False, Optional ArbitraryYes As Boolean = False) As Boolean
Shot
If Point(14676, 1116) = SluffRed Then
    EnemyInSight = True
End If
If EnemyInSight = False Then
    If ArbitraryYes Then
        If Point(14676, 1116) <> 0 Then
            EnemyInSight = True
        End If
    Else
        If ArbitraryNo Then
            EnemyInSight = False
        Else
            Snap FadeDark - Slept
            Shot
            If Point(14676, 1116) = SluffRed Then
                EnemyInSight = True
            End If
        End If
    End If
End If
End Function

Private Sub Snap(Ms As Long)
If Ms > 0 Then
    Slept = Slept + Ms
    Sleep Ms
    HasShot = False
End If
End Sub

Private Sub MakeHoney()
Do
k "2"
k " "
k "{enter}"
k "zr  "
Loop
End Sub

Private Sub QuickUse()
Do
k "{enter}"
k "3"
k "h"
Loop
End Sub

Private Function VDownstair() As Double
Dim VColour As Long
Dim Blur As Long
VColour = 9807520
Shot
Dim i As Double
For i = 0.442656606007002 To Pi * 2 + 0.148381351398575 Step 0.01
    For Blur = -30 To 30 Step 3
        If Point(ICenterX + Cos(i) * (VDownstairR + Blur), ICenterY + Sin(i) * (VDownstairR + Blur)) = VColour Then
            VDownstair = i
            Exit For
        End If
    Next Blur
Next i
If VDownstair = 0 Then
    VDownstair = 0.295518978702788
End If
End Function

Private Sub GoDown()
Dim Dis As Long
Dim RunTime As Integer
Dim WaitTime As Integer
Dis = Screen.Height * 0.7
Do
    If EnemyInSight Then
        k "p"
        k " "
        'Snap 300
    Else
        If Dis >= 50 Then
            CamOffSet VDownstair, Dis, True
            If RunTime <= 3 Then
                RunTime = RunTime + 1
            Else
                Dis = Dis * 0.5
            End If
        Else
            k " "
            WaitTime = 0
            RunTime = 0
            Dis = Screen.Height * 0.7
'            Do Until WaitTime >= 5000
'                snap 200
'                WaitTime = WaitTime + 200
'                If EnemyInSight Then
'                    WaitTime = 5000
'                End If
'            Loop
        End If
    End If
    'WaitTillSteady
    'Do Until IsSteady
    '    snap 500
    'Loop
Loop
End Sub

Private Sub CamOffSet(Angle As Double, ByVal Distance As Long, Optional Min As Boolean = False)
Debug.Assert Distance < Screen.Height
Angle = Angle + Pi
Dim MidX As Long, MidY As Long
Dim i As Single
If Min Then
    If Distance <= 100 Then
        Distance = 100
    End If
End If
Distance = Distance + 400
MidX = Screen.Width / 2
MidY = Screen.Height / 2
Zhi MidX - Distance * Cos(Angle) / 2, MidY - Distance * Sin(Angle) / 2
Dian 1
For i = 0 To 1 Step 0.05
    Zhi MidX + Distance * Cos(Angle) * (i - 0.5), MidY + Distance * Sin(Angle) * (i - 0.5)
    Sleep 10
Next i
Dian -1
End Sub

Private Function IsSteady() As Boolean
Shot
If 0 Then
    Dim CenterX As Long
    Dim CenterY As Long
    Dim R As Long
    CenterX = 1056
    CenterY = 1716
    R = 113
    Dim i As Single
    IsSteady = True
    For i = 0 To 2 * Pi Step 0.5
        If Point(CenterX + R * Cos(i), CenterY + R * Sin(i)) = vbWhite Then
            IsSteady = False
            Exit For
        End If
    Next i
End If

If 0 Then
    PicBox.Picture = LoadPicture(ShotPath)
    IsSteady = True
    Dim X As Long, Y As Long
    For X = 1 To Screen.Width Step 20
        For Y = 1 To Screen.Height Step 20
            If Me.Point(X, Y) <> PicBox.Point(X, Y) Then
                IsSteady = False
                Exit For
            End If
        Next Y
        If IsSteady = False Then
            Exit For
        End If
    Next X
    SavePicture Me.Image, ShotPath
End If


End Function

Private Sub Brake()
Dim C As Byte
Get #1, 1, C
If C <> 0 Then
    Unload Me
    End
End If
End Sub

Private Sub Hunt()
Dim WRS As Integer, WBS As Integer
Do
    Do Until EnemyInSight(True)
        'Debug.Assert WhereRedSluff = 0
        'Debug.Assert WhereBlueSluff = 0
        'Hunger = 0 '''''绝食
        If Hunger >= 800 Then
            k "{enter}"
            k "7"
            k "h"
            Hunger = 0
            Snap 700
        Else
            Me.Caption = Hunger
            If Wand > 6 Then
                k "2"
                Snap 300
                k "d"
                Wand = Wand - 1
                Snap 300
            Else
                k "v", 200
                Wand = 0 'NoWand
            End If
        End If
    Loop
    'Music 1
    'Debug.Assert WRS <> 0
    Do While EnemyInSight(True)
        k "p"
        k " "
        Snap 300
    Loop
    k " "
    If 0 Then
            k "c"
            WBS = WhereBlueSluff
            Do Until WBS = 0
                k " "
                Snap 300
                WBS = WhereBlueSluff
            Loop
            k "q"
            Music 13
            Snap 300
    End If
    If 0 Then
        k "2"
        k "s"
    End If
Loop
End Sub

Private Sub Hatch()
Dim i As Integer
    For i = 1 To 20
        k "2"
        k " "
        k "{enter}"
        k "d"
        k "5"
        k " "
    Next i
    k "{enter}"
    k "d"
    k "k"
    k "1"
    Unload Me
    End
End Sub

Private Sub TrainPet()
Do
    k "v"
Loop
End Sub

Private Function IsBusy() As Boolean
Shot
If Point(1128, 8676) = 9213843 Then
    IsBusy = True
End If
End Function

Private Sub HuntressHunt()
Dim WaitCount As Byte
Do
    If EnemyInSight(0, 1) Then
        If WaitCount >= 0 Then
            If SluffNotBlack Then
                k "m"
            Else
                k "2"
                k "2"
            End If
        Else
            WaitCount = WaitCount + 1
            If Not SluffNotBlack Then
                k "f"
            End If
        End If
    Else
        WaitCount = 0
        k "v"
    End If
Loop
End Sub

Private Sub Pitch()
Do
    k "1 ", 200
    k "{enter}", 200
    k "3", 200
    Snap 500
    k " "
    Snap 2000
    Brake
Loop
End Sub

Private Sub Super()
Dim i As Long, kk As Integer
Do
    'point(3648,1944~7920)=13731842
asd:
    For kk = 1 To 10
        Brake
        Snap 2000
        Shot
        For i = 1944 To 7920 Step 2
            If Point(3648, i) = 13731842 Then
                Exit For
            End If
        Next i
        If i < 7918 Then
            Zhi 3648, i
            Snap 500
            Dian
            kk = 10
        End If
    Next kk
    If i >= 7918 Then
        k "{down}"
        Snap 500
        k "{down}"
        Snap 500
        k "{down}"
        Snap 500
        GoTo asd
    End If
        
    Snap 2000
    'point(7764,5124)= 13731842
    Shot
    If Point(7764, 5124) = 13731842 Then
    'If 1 Then
        Zhi 7764, 5124
        Snap 500
        Dian
    Else
        Warn
    End If
    Snap 1000
    k "{pgdn}"
    Snap 1000
    k "{pgup}"
    DoEvents
Loop
End Sub

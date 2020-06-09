VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form AI 
   BorderStyle     =   0  'None
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   12
   ClientWidth     =   4656
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4656
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Timer Stopper 
      Index           =   0
      Interval        =   100
      Left            =   1080
      Top             =   360
   End
   Begin VB.Timer Reader 
      Left            =   3120
      Top             =   1680
   End
   Begin WMPLibCtl.WindowsMediaPlayer P 
      Height          =   852
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1092
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1926
      _cy             =   1503
   End
End
Attribute VB_Name = "AI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim Speaking As String
Public isReading As Boolean
Const PinYinMax As Integer = 395
Private Type ShengDiao
    PinYin(1 To 395) As String
End Type
Dim SD(1 To 4) As ShengDiao

Const ReaderSpeed As Integer = 333
Dim ReaderTP As Integer

Const Vol As Integer = 100

Dim Reading As Byte

Dim PreV As String

Private Sub Form_Load()
'把字库加载到缓存
For k = 1 To 4
    Open App.Path & "\r\" & k & ".txt" For Input As #k + 10
        For q = 1 To PinYinMax
            Input #k + 10, SD(k).PinYin(q)
        Next q
    Close #k + 10
Next k
With P(0)
    .settings.volume = 0
    .URL = App.Path & "\R\yy.wav"
End With
ReaderTP = Round(2000 / ReaderSpeed)
For k = 1 To ReaderTP - 1
    Load P(k)
    Load Stopper(k)
Next k
Reader.Interval = ReaderSpeed
End Sub

Private Sub Reader_Timer()
If Len(Speaking) = 0 Then
    Reader = False
    isReading = False
Else
    Dim Pst As Integer, N As String
    If Len(Speaking) >= 2 Then  '还有没有其他字？
        N = Mid(Speaking, 2, 1) '处理第三声连续问题
    Else
        N = "none"
    End If
    Pst = v(Left(Speaking, 1), N)
    If Pst = -1 Then GoTo Sk
    With P(Reading).Controls
        .currentPosition = Pst - 0.7 '
        .play
    End With
    P(Reading).settings.volume = Vol
    Stopper(Reading).Interval = 1900 '在下次接到任务之前停止。所以要搞一个小一点的值。
Sk: Reading = (Reading + 1) Mod ReaderTP
    Speaking = Right(Speaking, Len(Speaking) - 1)
End If
End Sub

Private Function v(C As String, N As String) As Integer
Dim k As Integer, q As Integer
If C = " " Then
    v = -1
Else
    GoTo NoExit
End If
GoTo ExtF
NoExit:
    For k = 1 To 4
        For q = 1 To PinYinMax
            If InStr(SD(k).PinYin(q), C) Then
                Select Case C
                    Case "难"
                        If PreV = "遇" Then
                            k = 4
                        End If
                End Select
                GoTo Found
            End If
        Next q
    Next k
    v = -1
    GoTo ExtF
Found: If N <> "none" Then  '后面有字
        If k = 3 Then   'C是第三声
            For ww = 1 To 4
                For w = 1 To PinYinMax
                    If InStr(SD(ww).PinYin(w), N) Then
                        GoTo FoundN
                    End If
                Next w
            Next ww
FoundN:     If ww = 3 Then
                k = 2
            End If
        End If
       End If
    v = ((k - 1) * PinYinMax + q) * 2
ExtF:
    PreV = C
End Function

Private Sub Stopper_Timer(Index As Integer)
Stopper(Index).Interval = 0
P(Index).Controls.pause
End Sub

Public Sub Speak(Words As String)
    Speaking = Speaking & Words
    Reader = True
    isReading = True
End Sub

Public Sub Mute()
    Speaking = " "
    isReading = False
    Reader = False
End Sub

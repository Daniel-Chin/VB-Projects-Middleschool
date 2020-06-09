VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "intelink"
   ClientHeight    =   5772
   ClientLeft      =   48
   ClientTop       =   1488
   ClientWidth     =   10884
   DrawStyle       =   2  'Dot
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   10884
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer WordMove 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   4320
      Top             =   4200
   End
   Begin VB.Timer Fade 
      Interval        =   40
      Left            =   3600
      Top             =   3600
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   5760
      Top             =   4920
   End
   Begin VB.TextBox Focusa 
      Height          =   495
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   -1900
      Width           =   1215
   End
   Begin VB.Label Busy 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "正在下载数据..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   3465
      TabIndex        =   5
      Top             =   4905
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Pad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pad"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   600
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Word 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Word"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   38.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   765
      Index           =   0
      Left            =   6600
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label ScorePad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   38.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   555
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Line TimeLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   19
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   5040
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "intelink"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   38.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   765
      Left            =   3878
      TabIndex        =   0
      Top             =   2505
      Width           =   3135
   End
   Begin VB.Menu Play 
      Caption         =   "开始游戏"
   End
   Begin VB.Menu Help 
      Caption         =   "帮助"
   End
   Begin VB.Menu Info1 
      Caption         =   "游戏作者是秦楠枫"
      Enabled         =   0   'False
   End
   Begin VB.Menu Info2 
      Caption         =   "内测版尚未向百度公司索取搜索引擎使用权，请勿将内测版外传"
      Enabled         =   0   'False
   End
   Begin VB.Menu Info3 
      Caption         =   "您在游戏中的行为将被记录。不要键入您的隐私内容。"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Gb As Integer
Dim Time As Single
Const MaxTime As Single = 10
Dim MoveStart As Integer
Dim Score As Single

Private Sub Fade_Timer()
If Pad.ForeColor >= Me.BackColor Then
    Fade = False
    Pad.ForeColor = Me.BackColor
Else
    Pad.ForeColor = Pad.ForeColor + RGB(10, 10, 10)
End If
End Sub

Private Sub Focusa_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        If Play.Enabled And Play.Caption = "开始游戏" Then
            Play_Click
        Else
            If Word(Gb) = "" Then
                PadLet "请输入一个单词"
            Else
                If Gb = 0 Then
                    BeBusy
                    If Web.Search(Word(Gb)) = 0 Or Web.Search(Word(Gb)) >= 50000000 Then
                        PadLet "单词过于常用"
                    Else
                        Title = Word(Gb)
                        GBCount
                        Timer = True
                        Time = MaxTime
                    End If
                Else
                    Dim k As Integer
                    For k = 0 To Gb - 1
                        If Word(k) = Word(Gb) Then
                            k = -1
                            Exit For
                        End If
                    Next k
                    If k = -1 Then
                        PadLet "单词重复"
                    Else
                        Dim Scoring As Single
                        BeBusy
                        Scoring = ToScore(Web.Check(Word(Gb), Title))
                        If Scoring = 0 Then
                            PadLet "相关度太低"
                            Time = Time / 2
                        Else
                            Time = Time + Scoring
                            If Time > MaxTime Then
                                Time = MaxTime
                            End If
                            Score = Score + Scoring
                            PadLet "+" & Format(Scoring, "0.0")
                            ScorePad = Format(Score, "0")
                            GBCount
                            'Title = Word(Int(Gb * Rnd))
                            Title = Word(Gb - 1)
                        End If
                    End If
                End If
            End If
        End If
    Case 65 To 90
        Word(Gb) = Word(Gb) & Chr(KeyCode + 32)
    Case 8
        If Len(Word(Gb)) >= 1 Then
            Word(Gb) = Left(Word(Gb), Len(Word(Gb)) - 1)
        End If
    Case 32
        Word(Gb) = Word(Gb) & " "
    Case 27
        Word(Gb) = ""
End Select
If Word(Gb).Left + Word(Gb).Width > Me.ScaleWidth Then
    Focusa_KeyDown 8, 0
End If
End Sub

Private Sub Form_Load()
Randomize
Web.Initialize
End Sub

Private Sub Help_Click()
MsgBox "游戏方法：输入与视图里给定单词相关的单词。" & Chr(10) & "单词相关度越高，分数就越高。" & Chr(10) & "当计时器归零时，游戏结束。" & Chr(10) & Chr(10) & _
"单词相关度的判定规则：" & Chr(10) & _
"设X=单词相关度" & Chr(10) & _
"设A=以“输入单词”为关键字，在百度搜索中的结果条目数" & Chr(10) & _
"设B=以“输入单词AND给定单词”为关键字，在百度搜索中的结果条目数" & Chr(10) & _
"设C=以“给定单词”为关键字，在百度搜索中的结果条目数" & Chr(10) & _
" X = log(B / sqr(A * C)) + 10. " & Chr(10) & "祝您游戏愉快。" & Chr(10) & Chr(10) & _
"您的输入单词将被记录，用以作为人工智能研发的材料。", _
vbOKOnly + vbInformation, "intelink帮助"
End Sub

Private Function ToScore(Checked As Integer) As Single
ToScore = Checked
End Function

Private Sub Play_Click()
If Play.Caption = "重新开始" Then
    Shell App.Path & "\" & App.EXEName, vbNormalFocus
    Unload Me
    End
Else
    Play.Enabled = False
    Help.Enabled = False
    With ScorePad
        .Caption = "0"
        .Left = (Me.ScaleWidth / 2 - .Width) / 2
        .Visible = True
    End With
    With Pad
        .Top = ScorePad.Height
        .Left = (Me.ScaleWidth / 2 - .Width) / 2
        .Caption = ""
        .Visible = True
    End With
    With Title
        .Caption = "请输入起点"
        .Left = (Me.ScaleWidth / 2 - .Width) / 2
    End With
    With TimeLine
        .Y2 = Me.ScaleHeight
        .BorderColor = vbRed
        .Visible = True
    End With
    With Word(0)
        .Top = (Me.ScaleHeight - .Height) / 2
        .Left = Me.ScaleWidth / 2
        .Caption = ""
        .Visible = True
    End With
End If
End Sub

Private Sub PadLet(Text As String)
Pad = Text
Pad.ForeColor = 0
Fade = True
End Sub

Private Sub GBCount()
WordMove = True
Gb = Gb + 1
Load Word(Gb)
Word(Gb - 1).BackStyle = 0
With Word(Gb)
    .Top = Word(Gb - 1).Top + Word(Gb - 1).Height
    .Left = Word(Gb - 1).Left
    .Caption = ""
    .BackStyle = 1
    .Visible = True
End With
End Sub

Private Sub Timer_Timer()
If Time <= 0 Then
    Si
    Timer = False
    WordMove = False
    Fade = False
Else
    Time = Time - Timer.Interval / 1000
    TimeLine.Y1 = Me.ScaleHeight * (1 - Time / MaxTime)
End If
End Sub

Private Sub WordMove_Timer()
If Word(Gb).Top <= (Me.ScaleHeight - Word(Gb).Height) / 2 Then
    Word(Gb).Top = (Me.ScaleHeight - Word(Gb).Height) / 2
    WordMove = False
Else
    Dim k As Integer
    For k = MoveStart To Gb
        With Word(k)
            .Top = .Top - .Height / 19
            If .Top + .Height < 0 Then
                MoveStart = k + 1
                .Visible = False
            End If
        End With
    Next k
End If
End Sub

Private Sub Si()
Info1.Visible = False
Info2.Visible = False
Info3.Visible = False
Help.Visible = False
Dim k As Integer
Dim PrX As Long, PrY As Long
For k = 0 To Gb - 1
    With Word(k)
        .FontSize = .FontSize / 1.9
        .Top = Int(Rnd * (Me.ScaleHeight - .Height))
        .Left = Int(Rnd * (Me.ScaleWidth - .Width))
        Line (PrX, PrY)-(.Left + .Width / 2, .Top + .Height / 2)
        PrX = .Left + .Width / 2
        PrY = .Top + .Height / 2
        .Visible = True
    End With
Next k
Word(Gb).Visible = False
Pad.Visible = False
TimeLine.Visible = False
Title.Visible = False
With ScorePad
    .FontSize = .FontSize * 2
    .BackStyle = 0
    .BorderStyle = 0
    .FontBold = True
    .ForeColor = vbRed
    .Top = (Me.ScaleHeight - .Height) / 2
    .Left = (Me.ScaleWidth - .Width) / 2
    .Caption = "分数：" & .Caption
End With
Play.Enabled = True
Play.Caption = "重新开始"
End Sub

Private Sub BeBusy()
Busy.Visible = True
DoEvents
Busy.Visible = False
End Sub

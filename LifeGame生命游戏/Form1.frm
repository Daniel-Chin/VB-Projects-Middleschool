VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "LifeGame"
   ClientHeight    =   6204
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   517
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   955
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Draft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1692
      Left            =   3000
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   171
      TabIndex        =   24
      Top             =   1320
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.FileListBox DirPad 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1464
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   2412
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   19
      Left            =   8040
      Top             =   3840
   End
   Begin VB.Frame Side 
      Caption         =   "Buttons"
      Height          =   6132
      Left            =   6720
      TabIndex        =   0
      Top             =   0
      Width           =   4692
      Begin VB.CommandButton Themer 
         Caption         =   "t&Heme"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   3240
         TabIndex        =   22
         Top             =   3240
         Width           =   1332
      End
      Begin VB.CommandButton JPGer 
         Caption         =   "&JPG"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   3240
         TabIndex        =   21
         Top             =   2640
         Width           =   1332
      End
      Begin VB.CommandButton Ruler 
         Caption         =   "&Rule"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1680
         TabIndex        =   15
         Top             =   3240
         Width           =   1332
      End
      Begin VB.CommandButton Importer 
         Caption         =   "i&Mport"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   3240
         TabIndex        =   19
         ToolTipText     =   "Also support BMP files"
         Top             =   2040
         Width           =   1332
      End
      Begin VB.CommandButton Exporter 
         Caption         =   "e&Xport"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   3240
         TabIndex        =   18
         Top             =   1440
         Width           =   1332
      End
      Begin VB.CommandButton Loader 
         Caption         =   "l&Oad"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   3240
         TabIndex        =   17
         Top             =   840
         Width           =   1332
      End
      Begin VB.CommandButton Saver 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   3240
         TabIndex        =   16
         Top             =   240
         Width           =   1332
      End
      Begin VB.CommandButton Intervaler 
         Caption         =   "&Interval"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   13.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1680
         TabIndex        =   14
         Top             =   2640
         Width           =   1332
      End
      Begin VB.CommandButton Slower 
         Caption         =   "&Kciuq"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1680
         TabIndex        =   13
         Top             =   2040
         Width           =   1332
      End
      Begin VB.CommandButton Quicker 
         Caption         =   "&Quick"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1680
         TabIndex        =   12
         Top             =   1440
         Width           =   1332
      End
      Begin VB.CommandButton Timerer 
         Caption         =   "&Timer"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1680
         TabIndex        =   11
         Top             =   840
         Width           =   1332
      End
      Begin VB.CommandButton Nexter 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   1332
      End
      Begin VB.CommandButton Eraser 
         Caption         =   "&Erase"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   120
         TabIndex        =   7
         Top             =   3240
         Width           =   1332
      End
      Begin VB.CommandButton Clearer 
         Caption         =   "&Blank"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1332
      End
      Begin VB.CommandButton Paint 
         Caption         =   "&Paint"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   1332
      End
      Begin VB.CommandButton Locker 
         Caption         =   "&Lock"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   1332
      End
      Begin VB.CommandButton Creater 
         Caption         =   "&Create"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1332
      End
      Begin VB.CommandButton ReGridder 
         Caption         =   "re&Grid"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1332
      End
      Begin VB.Label Prompt 
         Alignment       =   2  'Center
         Caption         =   "Prompt"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1452
         Left            =   120
         TabIndex        =   20
         Top             =   4560
         Width           =   4500
      End
   End
   Begin VB.Label Edge 
      BackColor       =   &H8000000D&
      Height          =   372
      Index           =   1
      Left            =   3960
      TabIndex        =   10
      Top             =   5400
      Width           =   972
   End
   Begin VB.Label Edge 
      BackColor       =   &H8000000D&
      Height          =   372
      Index           =   0
      Left            =   2280
      TabIndex        =   9
      Top             =   5280
      Width           =   972
   End
   Begin VB.Label Tile 
      Height          =   372
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Size As Byte
Dim Locked As Boolean
Dim Brushing As Boolean
Dim Erasing As Boolean
Dim LastDraw As Integer
Dim FpI As Integer
Dim DieWhenLow As Byte, DieWhenHigh As Byte, BirthWhenLow As Byte, BirthWhenHigh As Byte
Dim Slot() As Boolean
Dim Root As String

Private Sub Clearer_Click()
Dim k As Integer
For k = 1 To Size ^ 2
    Tile(k).BackColor = Me.BackColor
Next k
End Sub

Private Sub DirPad_DblClick()
DirPad.Visible = False
Dim k As Integer
Dim Readed As Boolean
Dim File As String, Lst As String
File = DirPad.FileName
Lst = Mid(File, InStr(File, ".") + 1, 4)
If Lst = "bmp" Or Lst = "BMP" Then
    With Draft
        .AutoSize = True
        .Picture = LoadPicture(Root & File)
        .AutoSize = False
        If .Width <> .Height Then
            MsgBox "The image must be a square.", vbCritical + vbOKOnly
            Exit Sub
        End If
        Prompt = "Matrix imported: " & File
        DoEvents
        ReSize .Height
        For k = 0 To .Height ^ 2 - 1
            If .Point(k Mod Size, k \ Size) <> 0 Then
                Tile(k + 1).BackColor = Me.ForeColor
            Else
                Tile(k + 1).BackColor = Me.BackColor
            End If
        Next k
    End With
    Readed = True
End If
If Lst = "smyx" Or Lst = "SMYX" Then
    Open Root & File For Binary As #1
    Prompt = "Matrix imported: " & File
    Dim C As Byte
    Get #1, 1, C
    ReSize C
    Get #1, 2, C
    DieWhenLow = C
    Get #1, 3, C
    DieWhenHigh = C
    Get #1, 4, C
    BirthWhenLow = C
    Get #1, 5, C
    BirthWhenHigh = C
    For k = 1 To Size ^ 2
        Get #1, 5 + k, C
        If C = 19 Then
            Tile(k).BackColor = Me.ForeColor
        Else
            Tile(k).BackColor = Me.BackColor
        End If
    Next k
    Close #1
    Readed = True
End If
If Not Readed Then
    MsgBox "Unknown file type.", vbCritical + vbOKOnly
    Prompt = "Matrix NOT imported"
End If
End Sub

Private Sub DirPad_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 27 'ESC
        DirPad.Visible = False
    Case 13 'Enter
        DirPad_DblClick
End Select
End Sub

Private Sub Eraser_Click()
Erasing = Not Erasing
If Erasing Then
    Eraser.Caption = "--&E--"
    If Brushing Then
        Paint_Click
    End If
Else
    Eraser.Caption = "&Erase"
End If
LastDraw = 0
End Sub

Private Sub Exporter_Click()
Dim File As String
File = InputBox("Input FileName. Current matrix will be saved as " & Root & "FileName.smyx", , Int(Rnd * 919))
If File = "" Then
    Prompt = "Matrix NOT exported"
    Exit Sub
Else
    File = File & ".smyx"
End If
If Dir(Root & File) <> "" Then
    If MsgBox("Such file already exists. Replace it?", vbYesNo + vbExclamation) = vbNo Then
        Prompt = "Matrix NOT exported"
        Exit Sub
    End If
End If
Open Root & File For Binary As #1
Dim C As Byte
C = Size
Put #1, 1, C
C = DieWhenLow
Put #1, 2, C
C = DieWhenHigh
Put #1, 3, C
C = BirthWhenLow
Put #1, 4, C
C = BirthWhenHigh
Put #1, 5, C
Dim k As Integer
For k = 1 To Size ^ 2
    If Tile(k).BackColor = Me.ForeColor Then
        C = 19
    Else
        C = 0
    End If
    Put #1, 5 + k, C
Next k
Close #1
Prompt = "Matrix exported as " & File
End Sub

Private Sub Importer_Click()
With DirPad
    .Visible = Not .Visible
    If .Visible Then
        .SetFocus
        .Refresh
    End If
End With
End Sub

Private Sub Intervaler_Click()
Dim NewInt As Integer, NewFpI As Integer
NewInt = Val(InputBox("Input new interval.", "Timer Setting", Timer.Interval))
If NewInt <= 0 Or NewInt > 19000 Then
    MsgBox "Invalid value.", vbExclamation + vbOKOnly
    Exit Sub
End If
NewFpI = Val(InputBox("Input new FramePerInterval.", "Timer Setting", FpI))
If NewFpI <= 0 Or NewFpI > 256 Then
    MsgBox "Invalid value.", vbExclamation + vbOKOnly
    Exit Sub
End If
With Timer
    .Interval = NewInt
    If .Enabled = True Then
        .Enabled = False
        .Enabled = True
    End If
    Prompt = "Interval: " & .Interval & Chr(10) & "FpI: " & FpI
End With
FpI = NewFpI
End Sub

Private Sub JPGer_Click()
Dim k As Integer, File As String
Draft.Width = Size
Draft.Height = Size
Draft.Cls
For k = 0 To Size ^ 2 - 1
    Draft.PSet (k Mod Size, k \ Size), IIf(Tile(k + 1).BackColor = Me.ForeColor, vbWhite, 0)
Next k
SavePicture Draft.Image, Root & "jpg.bmp"
Prompt = "BMP file generated"
End Sub

Private Sub Loader_Click()
Dim k As Integer
For k = 1 To Size ^ 2
    Tile(k).BackColor = IIf(Slot(k), Me.ForeColor, Me.BackColor)
Next k
Prompt = "Saving loaded"
End Sub

Private Sub Nexter_Click()
Dim k As Integer, LookK As Integer
Dim X As Integer, Y As Integer, Dens As Byte
For k = 1 To Size ^ 2
    Dens = 0
    For X = -1 To 1
        For Y = -1 To 1
            If Valu(k + X + Y * Size) And Not ((X = -1 And k Mod Size = 1) Or (X = 1 And k Mod Size = 0)) Then
                Dens = Dens + 1
            End If
        Next Y
    Next X
    With Tile(k)
        If .BackColor = Me.ForeColor Then
            If Dens <= DieWhenLow Or Dens >= DieWhenHigh Then
                .Tag = "0"
            End If
        Else
            If Dens >= BirthWhenHigh And Dens <= BirthWhenLow Then
                .Tag = "1"
            End If
        End If
    End With
Next k
For k = 1 To Size ^ 2
    With Tile(k)
        Select Case .Tag
            Case "0"
                .BackColor = Me.BackColor
            Case "1"
                .BackColor = Me.ForeColor
        End Select
        .Tag = ""
    End With
Next k
End Sub

Private Sub Paint_Click()
Brushing = Not Brushing
If Brushing Then
    Paint.Caption = "--&P--"
    If Erasing Then
        Eraser_Click
    End If
Else
    Paint.Caption = "&Paint"
End If
LastDraw = 0
End Sub

Private Sub Creater_Click()
Dim k As Integer
For k = 1 To Size ^ 2
    Tile(k).BackColor = IIf(Rnd > 0.5, Me.BackColor, Me.ForeColor)
Next k
Prompt = "Randome figure created"
End Sub

Private Sub Form_Load()
Randomize
Size = 64
Dim k As Integer
For k = 1 To Size ^ 2
    Load Tile(k)
Next k
ReGridder_Click
FpI = 1
DieWhenLow = 2
DieWhenHigh = 5
BirthWhenLow = 3
BirthWhenHigh = 3
Locker_Click
Prompt = "LifeGame by Daniel"
Root = App.Path & "\bin\"
DirPad.Path = Root
End Sub

Private Sub Form_Resize()
With Side
    .Left = Me.ScaleWidth - .Width
    .Height = Me.ScaleHeight
End With
With DirPad
    .Height = Me.ScaleHeight
    .Width = Side.Left
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Timer Then
    Cancel = 1
    Timerer_Click
Else
    End
End If
End Sub

Private Sub Locker_Click()
Locked = Not Locked
If Locked Then
    Locker.Caption = "Un&Lock"
    Prompt = "Click-Edit DISabled"
Else
    Locker.Caption = "&Lock"
    Prompt = "Click-Edit ENabled"
End If
End Sub

Private Sub Quicker_Click()
With Timer
    .Interval = Int(.Interval / 2) + 1
    If .Enabled = True Then
        .Enabled = False
        .Enabled = True
    End If
    Prompt = "Interval: " & .Interval / 1000 & "s"
End With
End Sub

Private Sub ReGridder_Click()
Me.BackColor = Me.ForeColor
Dim ChiCun As Long
ChiCun = Me.ScaleHeight
If ChiCun > Side.Left Then
    ChiCun = Side.Left
End If
ChiCun = Int(ChiCun / Size) * Size
Edge(0).Move 0, ChiCun, ChiCun, 19
Edge(1).Move ChiCun, 0, 19, ChiCun + 19
ChiCun = ChiCun / Size
Dim k As Integer, kk As Integer
For k = 1 To Size
    For kk = 1 To Size
        With Tile((k - 1) * Size + kk)
            .Width = ChiCun
            .Height = ChiCun
            .Left = (kk - 1) * ChiCun
            .Top = (k - 1) * ChiCun
            .Visible = True
        End With
    Next kk
    DoEvents
Next k
Me.BackColor = ReGridder.BackColor
End Sub

Private Sub Ruler_Click()
DieWhenLow = Val(InputBox("DieWhenLow", , DieWhenLow))
DieWhenHigh = Val(InputBox("DieWhenHigh", , DieWhenHigh))
BirthWhenLow = Val(InputBox("BirthWhenLow", , BirthWhenLow))
BirthWhenHigh = Val(InputBox("BirthWhenHigh", , BirthWhenHigh))
Prompt = "Rule Code: " & DieWhenLow & DieWhenHigh & BirthWhenLow & BirthWhenHigh
Dim NewSize As Byte
NewSize = Val(InputBox("Size", , Size))
If NewSize <> Size Then
    If MsgBox("Will reset. Make sure you have saved what you want. Press cancel to abort resetting.", vbOKCancel + vbQuestion, "Confirm reset") = vbOK Then
        ReSize NewSize
    Else
        Prompt = Prompt & Chr(10) & "Size not changed."
    End If
End If
End Sub

Private Sub Saver_Click()
If InputBox("Type ;' to confirm to save.", , "cancel") = ";'" Then
    ReDim Slot(1 To Size ^ 2) As Boolean
    Dim k As Integer
    For k = 1 To Size ^ 2
        Slot(k) = Tile(k).BackColor = Me.ForeColor
    Next k
    MsgBox "Saved.", vbOKOnly + vbInformation
    Prompt = "Saved"
Else
    Prompt = "Didn't save"
End If
End Sub

Private Sub Slower_Click()
If Timer.Interval <= 19000 Then
    Timer.Interval = Timer.Interval * 2
End If
Prompt = "Interval: " & Timer.Interval / 1000 & "s"
End Sub

Private Sub Tile_Click(Index As Integer)
If Not Locked Then
    With Tile(Index)
        If .BackColor = Me.ForeColor Then
            .BackColor = Me.BackColor
        Else
            .BackColor = Me.ForeColor
        End If
    End With
End If
End Sub

Private Sub Tile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If LastDraw <> Index Then
    If Brushing Then
        With Tile(Index)
            .BackColor = Me.ForeColor
        End With
        LastDraw = Index
    Else
        If Erasing Then
            With Tile(Index)
                .BackColor = Me.BackColor
            End With
            LastDraw = Index
        End If
    End If
End If
End Sub

Private Function Valu(Index As Integer) As Boolean
If Index >= 1 And Index <= Size ^ 2 Then
    If Tile(Index).BackColor = Me.ForeColor Then
        Valu = True
    End If
End If
End Function

Private Sub Timer_Timer()
Dim k As Integer
For k = 1 To FpI
    Nexter_Click
Next k
End Sub

Private Sub Timerer_Click()
Timer.Enabled = Not Timer.Enabled
If Timer Then
    Prompt = "Timer on"
Else
    Prompt = "Timer off"
End If
End Sub

Private Sub ReSize(NewSize As Byte)
Dim k As Integer
Clearer_Click
If NewSize > Size Then
    For k = Size ^ 2 + 1 To NewSize ^ 2
        Load Tile(k)
    Next k
Else
    For k = NewSize ^ 2 + 1 To Size ^ 2
        Unload Tile(k)
    Next k
End If
Size = NewSize
Prompt = Prompt & Chr(10) & "Size: " & Size
ReGridder_Click
End Sub

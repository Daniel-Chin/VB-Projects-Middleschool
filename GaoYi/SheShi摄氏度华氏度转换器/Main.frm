VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   5016
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8016
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Main"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5016
   ScaleWidth      =   8016
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Exiter 
      Enabled         =   0   'False
      Interval        =   36
      Left            =   4200
      Top             =   2880
   End
   Begin VB.PictureBox BG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   492
      ScaleWidth      =   1212
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Starter 
      Interval        =   36
      Left            =   3840
      Top             =   2280
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const Pi As Double = 3.14159265358979

Dim Tms As Integer
Const StartTms As Byte = 66
Dim StA As Single, StB As Single

Dim ShWid As Long, ShHei As Long
Dim TMD As Byte
Const ShTMD As Byte = 216

Dim ExitS As Long
Const ExitA As Long = 66

Dim Draging As Boolean
Dim DragX As Long, DragY As Long

Private Sub BG_Click()
AnNiu0.Show
End Sub

Private Sub BG_KeyPress(KeyAscii As Integer)
Form_KeyPress KeyAscii
End Sub

Private Sub Exiter_Timer()
ExitS = ExitS + ExitA
Top = Top + ExitS
If Top >= Screen.Width Then
    Unload Me
    End
End If
End Sub

Public Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 27
        If Exiter.Enabled Or Starter.Enabled Then
            Unload Me
            End
        Else
            '卸载所有悬浮控件
            Unload AnNiu0
            Unload AnNiu1
            Unload AnNiu2
            Unload AnNiu3
            Unload AnNiu4
            Me.PaintPicture LoadPicture(App.Path & "\bg0.jpg"), 0, 0, Me.Width, Me.Height
            BG.Visible = False
            ExitS = -500
            Exiter.Enabled = True
            Draging = False
        End If
End Select
End Sub

Private Sub Form_Load()
Starter.Interval = 26
ShWid = Width
ShHei = Height
'Me.Tag = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)  '-1为置前，-2
SetWindowLongA hwnd, -20, 524288
TMD = 66
SetTMD
With Me
    .Top = Screen.Height / 2
    .Height = 0
    .Left = Screen.Width / 2
    .Width = 0
    .BackColor = 0
    BG.Visible = False
    BG.Width = ShWid
    BG.Height = ShHei
    BG.PaintPicture LoadPicture(App.Path & "\bg.jpg"), 0, 0, ShWid, ShHei
End With
StA = -2 / StartTms ^ 2
StB = 3 / StartTms
Show
End Sub

Public Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Not Starter.Enabled And Not Exiter.Enabled Then
    DragX = X
    DragY = Y
    Draging = True
    TMD = 100
    SetTMD
End If
End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RCl
If Draging Then
    Move Left + X - DragX, Top + Y - DragY
    '移动悬浮控件
    With AnNiu0
        .Move .Left + X - DragX, .Top + Y - DragY
    End With
    With AnNiu1
        .Move .Left + X - DragX, .Top + Y - DragY
    End With
    With AnNiu2
        .Move .Left + X - DragX, .Top + Y - DragY
    End With
    With AnNiu3
        .Move .Left + X - DragX, .Top + Y - DragY
    End With
    With AnNiu4
        .Move .Left + X - DragX, .Top + Y - DragY
    End With
End If
End Sub

Public Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Not Starter.Enabled And Not Exiter.Enabled Then
    Draging = False
    TMD = ShTMD
    SetTMD
End If
End Sub

Private Sub BG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub BG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub BG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub Starter_Timer()
Tms = Tms + 1
If Tms = StartTms Then
    TMD = ShTMD
    SetTMD
    Me.Height = ShHei
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Width = ShWid
    Me.Left = (Screen.Width - Me.Width) / 2
    BG.Visible = True
    Starter.Enabled = False
    '显示悬浮控件
    With AnNiu0
        .Show
        .Left = Left + 1200
        .Top = Top + 900
    End With
    With AnNiu1
        .Show
        .Left = Left + 1850
        .Top = Top + 2500
    End With
    With AnNiu2
        .Show
        .Left = Left + 1850
        .Top = Top + 3700
    End With
    With AnNiu3
        .Show
        .Left = Left + 4200
        .Top = Top + 2500
    End With
    With AnNiu4
        .Show
        .Left = Left + 4200
        .Top = Top + 3700
    End With
Else
    TMD = TMD + 2
    SetTMD
    With Me
        .Height = ShHei * (StA * Tms ^ 2 + StB * Tms)
        .Width = ShWid * (StA * Tms ^ 2 + StB * Tms)
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
        If Tms <= StartTms / 4 * 3 Then
            .BackColor = .BackColor + RGB(2, 2, 0)
        Else
            .BackColor = .BackColor - RGB(2, 2, 0) * 3
        End If
    End With
End If
End Sub

Private Sub SetTMD()
SetLayeredWindowAttributes hwnd, 0, TMD, 2
End Sub

Public Sub RCl()
AnNiu1.Label1.BackColor = 74565
AnNiu1.Label1.ForeColor = vbWhite
AnNiu2.Label1.BackColor = 74565
AnNiu2.Label1.ForeColor = vbWhite
End Sub

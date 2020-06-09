VERSION 5.00
Begin VB.Form Lv13Hider 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5796
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8700
   Icon            =   "Lv13Hider.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5796
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Droper 
      Interval        =   36
      Left            =   4440
      Top             =   3360
   End
   Begin VB.Timer Main 
      Left            =   5160
      Top             =   2280
   End
   Begin VB.Timer Loader 
      Left            =   5520
      Top             =   3120
   End
   Begin VB.Label WH 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   500
   End
   Begin VB.Label Y 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5270
      TabIndex        =   1
      Top             =   380
      Width           =   260
   End
   Begin VB.Line JG 
      BorderColor     =   &H000000FF&
      BorderWidth     =   6
      X1              =   2880
      X2              =   4800
      Y1              =   2360
      Y2              =   2360
   End
   Begin VB.Label Z 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2740
      TabIndex        =   0
      Top             =   600
      Width           =   260
   End
End
Attribute VB_Name = "Lv13Hider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DYC As Boolean
Public Tms As Byte

Public Draging As Boolean, DragY As Long
Public WYY As Long, PreT As Long

Private Sub Droper_Timer()
WH.Top = WH.Top + 36
If WH.Top >= 1656 Then
    Lv13.Si True, "你被砸死了！"
End If
End Sub

Private Sub Form_Load()
Loader.Interval = 1
Show
DYC = -1
End Sub

Private Sub Form_LostFocus()
Me.SetFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Draging = True
DragY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Draging And Y > DragY Then
    If WYY <= 2666 Then
        WYY = WYY + (Y - DragY) / (36 - 35 / 2666 * WYY)
    Else
        WYY = WYY + (Y - DragY) / 1
    End If
    DragY = Y
    If WYY >= 5666 Then
        Unload Me
    End If
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Draging = False
End Sub

Private Sub Loader_Timer() '伪
Main.Interval = 1000
With Lv13
    Top = .Top + 666 + WYY
    DragY = DragY - Top + PreT
    Left = .Left
    Height = .Height - 666 - WYY
    Width = .Width
    PreT = Top
    .Hd.Top = WYY
End With
If DYC Then
    DYC = 0
    Tms = 0
    JG.Visible = False
    WH.Visible = False
    PaintPicture LoadPicture(App.Path & "\chat.jpg"), 0, 0, Width, Height
    WH.Visible = True
End If
End Sub

Private Sub Main_Timer()
Tms = Tms + 1
Select Case Tms
    Case 2
        Y = "嘿，兄弟！"
    Case 4
        Y = ""
        Z = "什么事？"
    Case 6
        Y = "你觉得金钱重要还是女人重要？"
        Z = ""
    Case 8
        Y = ""
        Z = "女人。"
    Case 11
        Y = "明明是金钱重要！"
        Z = ""
    Case 14
        Y = ""
        Z = "我不这么认为。"
    Case 17
        Y = "你完了。"
        Z = ""
    Case 20
        Y = ""
        JG.Visible = True
    Case 23
        Lv13.Si True, "你被激光枪射死了！"
End Select
End Sub

Private Sub WH_Click()
Droper = True
End Sub

Private Sub Z_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Z_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Z_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub Y_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Y_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Y_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseUp Button, Shift, X, Y
End Sub


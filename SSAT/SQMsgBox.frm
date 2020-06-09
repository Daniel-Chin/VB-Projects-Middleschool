VERSION 5.00
Begin VB.Form SQMsgBox 
   BackColor       =   &H00CC0000&
   BorderStyle     =   0  'None
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   Icon            =   "SQMsgBox.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Loader 
      Interval        =   1
      Left            =   3600
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   1440
   End
   Begin VB.Label TT 
      Alignment       =   2  'Center
      BackColor       =   &H00CC0000&
      Caption         =   "恭喜你获得了一个神器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   8172
   End
   Begin VB.Label ZCD 
      Alignment       =   2  'Center
      BackColor       =   &H00CC0000&
      Caption         =   "恭喜你获得了一个神器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   8172
   End
   Begin VB.Label PM 
      Alignment       =   2  'Center
      BackColor       =   &H00CC0000&
      Caption         =   "恭喜你获得了一个神器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   8172
   End
   Begin VB.Label Bt 
      Alignment       =   2  'Center
      BackColor       =   &H00CC0000&
      Caption         =   "恭喜你刚刚获得了一个神器"
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
      Height          =   852
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   8172
   End
End
Attribute VB_Name = "SQMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public SQ_ID As Integer
Dim TMD As Integer
Dim V As Integer

Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Sub Bt_Click()
Timer1 = True
End Sub

Private Sub TT_Click()
Timer1 = True
End Sub

Private Sub PM_Click()
Timer1 = True
End Sub

Private Sub ZCD_Click()
Timer1 = True
End Sub

Private Sub Form_Click()
Timer1 = True
End Sub

Private Sub Loader_Timer()
Dim CP()
Tag = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
If GetSetting("iGlope", "woqu", "sq" & SQ_ID, 0) = 1 Then
    End
Else
    Timer1.Interval = 36
    V = 16
    SetWindowLongA hwnd, -20, 524288
    SetLayeredWindowAttributes hwnd, 0, 0, 2
    CP = Array("完成游戏基本关卡", "右击主菜单标题6次", "让游戏溢出", "第十三关向左走", "使用秘籍", "集齐全部六个神器")
    Me.Left = 0
    Me.Width = Screen.Width
    PM.Width = Me.Width
    ZCD.Width = Me.Width
    TT.Width = Me.Width
    Bt.Width = Me.Width
    PM = Chr(34) & CP(SQ_ID - 1) & Chr(34)
    PM.ForeColor = &H6CFFFF
    ZCD = "快去主菜单看看吧！"
    ZCD.FontItalic = True
    TT = "我去冒险 提醒您"
    TT.FontBold = True
    SaveSetting "iGlope", "woqu", "sq" & SQ_ID, "1"
End If
Loader = False
End Sub

Private Sub Timer1_Timer()
If Timer1.Interval = 3666 Then
    Timer1.Interval = 36
    Timer1 = False
    Exit Sub
End If
If TMD < 255 - V And TMD > -V Then
    TMD = TMD + V
    SetLayeredWindowAttributes hwnd, 0, TMD, 2
Else
    If V < 0 Then End
    Timer1.Interval = 3666
    V = -V
    TMD = TMD + V
End If
End Sub

VERSION 5.00
Begin VB.Form FrmWndProc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   375
      Left            =   405
      TabIndex        =   1
      Top             =   450
      Width           =   3210
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5130
      Top             =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "在这里滚动滚轮看看"
      Top             =   135
      Width           =   3210
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   945
      Width           =   2310
   End
End
Attribute VB_Name = "FrmWndProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
'**说    明：紫水晶工作室 http://www.m5home.com/
'**创 建 人：马大哈
'**日    期：2005年04月13日
'**描    述：一个使用子类技术得到鼠标滚轮状态的简单例子
'*************************************************************************
Option Explicit

Dim St As Long

Private Sub Command1_Click()
                                '这里是对Text1进行子类化处理
If St = -1 Then

    PrevWndProc = SetWindowLong(Text1.Hwnd, GWL_WNDPROC, AddressOf SubWndProc)
    Command1.Caption = "停止"
    St = 1
    Me.Caption = "子类处理状态!"
     
Else

    SetWindowLong Text1.Hwnd, GWL_WNDPROC, PrevWndProc
    Command1.Caption = "开始"
    St = -1
    Me.Caption = "正常状态"
    
End If

End Sub

Private Sub Form_Load()

St = -1
Me.Caption = "准备完毕"

End Sub

Private Sub Form_Unload(Cancel As Integer)

If St <> -1 Then
    SetWindowLong Text1.Hwnd, GWL_WNDPROC, PrevWndProc
End If

End Sub

Private Sub Timer1_Timer()

Select Case MouseW
    Case 1
        Label1.Caption = "向上滚动"
    Case -1
        Label1.Caption = "向下滚动"
    Case Else
        Label1.Caption = "滚轮静止"
End Select

MouseW = 0

End Sub
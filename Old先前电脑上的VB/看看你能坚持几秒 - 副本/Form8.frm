VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   Caption         =   "本游戏由麻瓜传记出品"
   ClientHeight    =   6324
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11016
   LinkTopic       =   "Form8"
   ScaleHeight     =   6324
   ScaleWidth      =   11016
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   9480
      Top             =   1680
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The Tale of Two Maguas"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   22.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   612
      Left            =   3120
      TabIndex        =   1
      Top             =   3480
      Width           =   5052
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "麻瓜传记"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1692
      Left            =   4680
      TabIndex        =   0
      Top             =   1800
      Width           =   1812
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Form8
Form7.Visible = True
End Sub

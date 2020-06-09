VERSION 5.00
Begin VB.Form Starter 
   Caption         =   "语文学霸是怎样炼成的"
   ClientHeight    =   6324
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10644
   LinkTopic       =   "Form1"
   ScaleHeight     =   6324
   ScaleWidth      =   10644
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton YLW 
      Caption         =   "议论文"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   6000
      TabIndex        =   2
      Top             =   4440
      Width           =   2412
   End
   Begin VB.CommandButton SMW 
      Caption         =   "说明文"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   3960
      TabIndex        =   1
      Top             =   2040
      Width           =   2412
   End
   Begin VB.CommandButton JXW 
      Caption         =   "记叙文"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2412
   End
End
Attribute VB_Name = "Starter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ZhongXin As String, JiXuDuiXiang As String

Private Sub JXW_Click()
MsgBox "读文章，理清文章层次。画出过渡句。圈画关键的描述性词语，尽量找出中心，作者观点或人物品格。如果试题中有选择题，请先把各个选项读一遍，可能会有启发。", vbOKOnly, "记叙文"
If MsgBox("本文是叙事还是志人？叙事选择OK，志人选择Cancel。", vbOKCancel + vbQuestion, "记叙文") = vbOK Then
    '叙事
Else
    '志人
End If
End Sub


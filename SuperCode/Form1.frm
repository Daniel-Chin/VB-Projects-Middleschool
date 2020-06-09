VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "超级密码"
   ClientHeight    =   6228
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11388
   LinkTopic       =   "Form1"
   ScaleHeight     =   6228
   ScaleWidth      =   11388
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton bC 
      Caption         =   "Create"
      Height          =   492
      Index           =   2
      Left            =   10320
      TabIndex        =   11
      ToolTipText     =   "Create a randome key. "
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton bC 
      Caption         =   "Create"
      Height          =   492
      Index           =   1
      Left            =   10320
      TabIndex        =   7
      ToolTipText     =   "Create a randome key. "
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton bC 
      Caption         =   "&Create"
      Height          =   492
      Index           =   0
      Left            =   10320
      TabIndex        =   3
      ToolTipText     =   "Create a randome key. "
      Top             =   480
      Width           =   972
   End
   Begin SuperCode.Jdt Jdt 
      Height          =   492
      Index           =   0
      Left            =   3228
      Top             =   3480
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   868
   End
   Begin VB.CommandButton bD 
      Caption         =   "Decript"
      Height          =   492
      Index           =   2
      Left            =   9228
      TabIndex        =   10
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton bE 
      Caption         =   "Encript"
      Height          =   492
      Index           =   2
      Left            =   8148
      TabIndex        =   9
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton bD 
      Caption         =   "Decript"
      Height          =   492
      Index           =   1
      Left            =   9228
      TabIndex        =   6
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton bE 
      Caption         =   "Encript"
      Height          =   492
      Index           =   1
      Left            =   8148
      TabIndex        =   5
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton bD 
      Caption         =   "&Decript"
      Height          =   492
      Index           =   0
      Left            =   9228
      TabIndex        =   2
      Top             =   480
      Width           =   972
   End
   Begin VB.CommandButton bE 
      Caption         =   "&Encript"
      Height          =   492
      Index           =   0
      Left            =   8148
      TabIndex        =   1
      Top             =   480
      Width           =   972
   End
   Begin VB.TextBox Rt 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   7800
   End
   Begin VB.TextBox Rt 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   7800
   End
   Begin VB.TextBox Rt 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7800
   End
   Begin SuperCode.Jdt Jdt 
      Height          =   492
      Index           =   1
      Left            =   3228
      Top             =   4160
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   868
   End
   Begin SuperCode.Jdt Jdt 
      Height          =   492
      Index           =   2
      Left            =   3228
      Top             =   4840
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   868
   End
   Begin SuperCode.Jdt Jdt 
      Height          =   492
      Index           =   3
      Left            =   3228
      Top             =   5520
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   868
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim k As Long
Dim Key1(0 To 255) As Byte, Key2(0 To 255) As Byte

Private Sub bC_Click(Index As Integer)
'生成随机密钥
CrRnd
Dim Root As String
Root = Rt(Index)
Open Root & "\key" For Binary As #1
For k = 1 To 256
    Put #1, k, Key2(k - 1)
Next k
Close #1
MsgBox "Creation done."
End Sub

Private Sub bD_Click(Index As Integer)
'解密
Dim Root As String
Root = Rt(Index)
Open Root & "\key" For Binary As #1
If LOF(1) <> 256 Then MsgBox "debug"
For k = 1 To 256
    Get #1, k, Key2(k - 1)
Next k
Close #1
Open Root & "\code" For Binary As #3
On Error Resume Next
Kill Root & "\text.txt"
Open Root & "\text.txt" For Binary As #2
Dim tB As Byte
Dim p As Byte
For k = 1 To LOF(3) Step 257
    Get #3, k, tB
    Put #2, (k \ 257) + 1, AKey2(tB)
    For p = 0 To 255
        Get #3, k + 1 + p, tB
        Key1(p) = AKey2(tB)
    Next p
    For p = 0 To 255
        Key2(p) = Key1(p)
    Next p
    If Rnd() <= 0.1 Then
        RenewJdt k / LOF(3)
    End If
Next k
Close #2
Close #3
For k = 0 To 3
    Jdt(k).RenewProgress 1
Next k
MsgBox "Decription done."
End Sub

Private Sub bE_Click(Index As Integer)
'加密
Dim Root As String
Root = Rt(Index)
Open Root & "\key" For Binary As #1
If LOF(1) <> 256 Then MsgBox "debug"
For k = 1 To 256
    Get #1, k, Key1(k - 1)
Next k
Close #1
Open Root & "\text.txt" For Binary As #2
On Error Resume Next
Kill Root & "\code"
Open Root & "\code" For Binary As #3
Dim tB As Byte
Dim p As Integer
For k = 1 To LOF(2)
    CrRnd
    Get #2, k, tB
    Put #3, (k - 1) * 257 + 1, Key1(tB)
    For p = 0 To 255
        Put #3, (k - 1) * 257 + 2 + p, Key1(Key2(p))
    Next p
    For p = 0 To 255
        Key1(p) = Key2(p)
    Next p
    If Rnd() <= 0.1 Then
        RenewJdt k / LOF(2)
    End If
Next k
Close #2
Close #3
For k = 0 To 3
    Jdt(k).RenewProgress 1
Next k
MsgBox "Encription done."
End Sub

Private Sub Form_Load()
Randomize
Me.BackColor = RGB(200, 200, 200)
For k = 0 To 3
    With Jdt(k)
        .RenewBackColor 255, 255, 255
        .RenewEmptyColor 0, 0, 0
        .RenewFilledColor 0, 255, 0
        .RenewBars 6
    End With
Next k
For k = 0 To 2
    Rt(k).Text = App.Path & "\root\"
Next k
End Sub

Private Sub RenewJdt(JD As Double)
Dim jDg As Integer
jDg = Int(JD * 1296)
Jdt(0).RenewProgress jDg / 6 ^ 4
Jdt(1).RenewProgress (jDg Mod 6 ^ 3) / 6 ^ 3
Jdt(2).RenewProgress (jDg Mod 6 ^ 2) / 6 ^ 2
Jdt(3).RenewProgress (jDg Mod 6) / 6
DoEvents
End Sub

Private Sub Rt_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii
    Case 1
        Rt(Index).SelStart = 0
        Rt(Index).SelLength = Len(Rt(Index))
End Select
End Sub

Private Sub CrRnd()
Dim CP(0 To 255) As Byte
Dim p As Integer
For p = 0 To 255
    CP(p) = p
Next p
Dim rD As Byte
For p = 0 To 255
    rD = Int(Rnd() * (256 - p))
    Key2(CP(rD)) = p
    CP(rD) = CP(255 - p)
Next p
End Sub

Private Function AKey2(Index As Byte) As Integer
For AKey2 = 0 To 255
    If Key2(AKey2) = Index Then
        Exit For
    End If
Next AKey2
End Function

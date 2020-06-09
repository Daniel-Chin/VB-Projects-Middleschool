VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Ğ³Òô×ª»»"
   ClientHeight    =   6432
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9672
   BeginProperty Font 
      Name            =   "ËÎÌå"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6432
   ScaleWidth      =   9672
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "ºÇ"
      Height          =   972
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   492
   End
   Begin VB.CommandButton C 
      Caption         =   "×ª»»"
      Height          =   732
      Left            =   4170
      MouseIcon       =   "Form1.frx":044A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2760
      Width           =   1332
   End
   Begin VB.TextBox R 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2652
      Left            =   870
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3450
      Width           =   7932
   End
   Begin VB.TextBox T 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2652
      Left            =   870
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   330
      Width           =   7932
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s(2 To 20) As String
Dim p(2 To 20) As String

Private Sub C_Click()
R = ""
Dim i As Integer, q As Byte
Dim b As Boolean
For i = 1 To Len(T)
    For q = 2 To 20
        If InStr(s(q), Mid(T, i, 1)) <> 0 Then
            R = R & p(q)
            b = True
            GoTo TG
        End If
    Next q
TG:    If b = False Then
        R = R & Mid(T, i, 1)
    End If
    b = False
Next i
End Sub

Private Sub Form_Load()
For i = 2 To 9
    p(i) = i
Next i
p(10) = "0"
p(11) = "B"
p(12) = "C"
p(13) = "D"
p(14) = "E"
p(15) = "G"
p(16) = "I"
p(17) = "O"
p(18) = "P"
p(19) = "R"
p(20) = "T"
s(2) = "¶ø¶ş¶ú¶ù¶ü¶û·¡"
s(3) = "ÈıÉ¢É¡ÈşâÌôÖ"
s(4) = "ËÄËÀË¿ËºËÆË½Ë»Ë¼ËÂË¾Ë¹ËÃËÈËÇËÁËÅÊ³²ŞïÈ"
s(5) = "ÎŞÎåÎİÎïÎğÎíÎäÎóÎâÎæÎÛÎòÎèÎÙÎìÎñÎØÎé"
s(6) = "ÁùÁ÷ÁôÁõÁøÁïÁòÁöÁñÁğÁó"
s(7) = "ÆğÆäÆßÆøÆÚÆëÆ÷ÆŞÆïÆûÆåÆæÆÛÆáÆôÆİÆâÆñÆöÆúÆüÆàçùÆóç²ç÷ÆòÆõÆçÆíÆù"
s(8) = "°Ñ°Ë°É°Ö°Î°Õ°Ï°Í°Ô°Ç°Å°È°ĞôÎ"
s(9) = "¾Í¾Å¾Æ¾É¾Ã¾¾¾È¾À¾Ë¾¿¾Â"
s(10) = "ÁÖÁÙÁÜÁÚÁ×ÁÛÁŞÁßÁÕÁİÁíÁîÁìÁãÁåÁáÁéÁëÁäÁèÁêÁâÁæÁçÀâÜßñöñö"
s(11) = "±È±Ê±Õ±Ç±Ì±Ø±Ü±Æ±Ï±Û±Ë±É±Ú±Í±Ò±×±Ù±Î±Ğ±Ó±Ö±İ±Ñ±ÔÃØÃÚåöèµ"
s(12) = "Î÷Ï´Ï¸ÎüÏ·ÏµÏ²Ï¯Ï¡ÏªÏ¨ÎıÏ¥Ï¢Ï®Ï§Ï°ÎûÏ¦"
s(13) = "µØµÚµ×µÍµÜµÖµÎµÏµİµÙµÛµÕµĞµŞµÌµÓ"
s(14) = "Ò»ÒÔÒÑÒÚÒÂÒÆÒ×ÒÀÒåÒéÒ½ÒÒÒÇÒàÒÎÒæÒĞÒÌÒíÒëÒÁÒâÒãÒÕÒËÒáÒÓÒìÒÉÒÅÒêÒäÒÛÒİÒÙÒÖÒ¼"
s(15) = "¼¸¼°¼±¼È¼´»ú¼¦»ı¼Ç¼¶¼«¼§»ù¼Ä¼Í¼ª¼½¼Æ¼·¼º¼¾Ïµ¼¤¼¹¼Ê¼¼¼¯"
s(16) = "°®°«°¤°¥°­°©°¬°¦°§°ª"
s(17) = "Å¼Å»Å·ÅºÅ¸ÇøÅ½Å¹ñîÚ©"
s(18) = "ÅúÆ¤ÅûÆ¥Åü±ÙÅ÷Æ¨ÅúÆ§Æ£Æ¦ÅùÅıÅşÆ¡Æ©ÅøØ§"
s(19) = "°¡°¢ß¹"
s(20) = "ÌáÌæÌåÌâÌßÌãÌêÌŞÌİÌàÌäÌéÌçÌëç¾"
End Sub

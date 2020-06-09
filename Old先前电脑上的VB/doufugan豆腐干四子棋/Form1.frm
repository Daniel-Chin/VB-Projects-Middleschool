VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "豆腐干四子棋"
   ClientHeight    =   5616
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10416
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5616
   ScaleWidth      =   10416
   StartUpPosition =   3  '窗口缺省
   Begin VB.Shape Qi 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   492
      Index           =   0
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   3120
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label bT 
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   4680
      TabIndex        =   1
      Top             =   2640
      Width           =   972
   End
   Begin VB.Label Tray 
      BackColor       =   &H00FFFFC0&
      Height          =   492
      Left            =   3000
      TabIndex        =   0
      Top             =   1800
      Width           =   1092
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFFFF&
      Index           =   122
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sH As Integer, GeSize As Integer
Const Size As Byte = 3
Dim Draging As Boolean, DragX As Integer, DragY As Integer
Dim gE(0 To 30, 0 To 30) As Byte
Dim Fang As Byte
Dim S As Integer

Private Sub bT_Click()
Form_KeyUp 27, 0
End Sub

Private Sub bT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bT.BackColor = vbRed
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
        If MsgBox("Exit?", vbYesNo + vbQuestion, "Exit") = vbYes Then
            End
        End If
End Select
End Sub

Private Sub Form_Load()
'msgbox "请把任务栏拖到屏幕右侧。"
sH = Screen.Height
With Me
    .Height = sH
    .Width = sH + 350
    .Left = (Screen.Width - .Width) / 2
    .Top = 0
End With
With Tray
    .Top = 0
    .Left = sH
    .Height = sH - 350
    .Width = 350
    .FontSize = 15
    .Caption = "豆腐干五子棋"
End With
With bT
    .Left = sH
    .Top = sH - 350
    .Width = 350
    .Height = 350
    .FontSize = 15
    .Caption = "X"
End With
GeSize = sH / Size / 10
For i = 0 To Size * 10
    Load L(i)
    With L(i)
        .X1 = i * GeSize
        .X2 = .X1
        .Y1 = 0
        .Y2 = sH
        If i Mod 5 = 0 Then
            .BorderColor = vbGreen
            .BorderWidth = 2
        Else
            .BorderColor = vbWhite
            .BorderWidth = 1
        End If
        .Visible = True
    End With
    Load L(i + Size * 10 + 1)
    With L(i + Size * 10 + 1)
        .X1 = 0
        .X2 = sH
        .Y1 = i * GeSize
        .Y2 = .Y1
        If i Mod 5 = 0 Then
            .BorderColor = vbGreen
            .BorderWidth = 2
        Else
            .BorderColor = vbWhite
            .BorderWidth = 1
        End If
        .Visible = True
    End With
Next i
Fang = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bT.BackColor = &HC0&
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
    Case 1  '左键
        If gE(P(X), P(Y)) = 0 Then
            S = S + 1
            gE(P(X), P(Y)) = Fang
            Load Qi(S)
            With Qi(S)
                .Width = GeSize / 4 * 3
                .Height = .Width
                .Top = P(Y) * GeSize - .Height / 2
                .Left = P(X) * GeSize - .Width / 2
                .FillColor = IIf(Fang = 1, vbWhite, 0)
                .Visible = True
            End With
            
            Fang = 3 - Fang
            
        End If
End Select
End Sub

Private Sub Tray_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Draging = True
DragX = X
DragY = Y
End Sub

Private Sub Tray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Draging Then
    With Me
        .Left = .Left + X - DragX
        .Top = .Top + Y - DragY
    End With
End If
bT.BackColor = &HC0&
End Sub

Private Sub Tray_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Draging = False
End Sub

Private Function P(Length As Single) As Byte
    P = (Length + GeSize / 2) \ GeSize
End Function

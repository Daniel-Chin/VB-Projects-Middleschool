VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6108
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11076
   LinkTopic       =   "Form1"
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   923
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   5040
      Top             =   2880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Pi As Double = 3.14159265358979
Const Rotation As Single = 0.1
Const Speed As Long = 6
Const RotatePoss As Single = 0.19
Const CopyPoss As Single = 0.38
Const Size As Integer = 19
Const Interval As Integer = 1

Dim Map(0 To 9999) As Integer
Dim MapTop As Integer

Private Type Daniels
    Exist As Boolean
    X As Long
    Y As Long
    Rad As Double
    F As Integer
    R As Byte
    G As Byte
    B As Byte
End Type

Dim Daniel(0 To 9999) As Daniels

Private Sub Form_Click()
Cls
End Sub

Private Sub Form_Load()
Randomize
Timer1.Interval = Interval
With Daniel(0)
    .Exist = True
    .X = 0
    .Y = 0
    .Rad = Pi / 4
    .F = -1
    .R = 15
    .G = 15
    .B = 15
End With
Map(0) = 0
MapTop = 1
Me.DrawWidth = 2
End Sub

Private Sub Timer1_Timer()
Dim K As Integer
Dim OldX As Long, OldY As Long
K = 0
Do Until K >= MapTop
    With Daniel(Map(K))
        Debug.Assert .Exist
        'Debug.Assert Map(K) <> 10
        'Move
        If Rnd() < RotatePoss Then
            Debug.Assert .F <> 0
            .F = -.F
        End If
        .Rad = .Rad + .F * Rotation
        OldX = .X
        OldY = .Y
        .X = .X + Speed * Cos(.Rad)
        .Y = .Y + Speed * Sin(.Rad)
        Line (.X, .Y)-(OldX, OldY)
        'Bounce
        If .X < 0 Then
            .Rad = Pi - .Rad
            .X = Speed
        End If
        If .X > Me.ScaleWidth Then
            .Rad = Pi - .Rad
            .X = Me.ScaleWidth - Speed
        End If
        If .Y > Me.ScaleHeight Then
            .Rad = -.Rad
            .Y = Me.ScaleHeight - Speed
        End If
        If .Y < 0 Then
            .Rad = -.Rad
            .Y = Speed
        End If
        'Paint
        Me.ForeColor = RGB(.R, .G, .B) * 17
        'Die
        Debug.Assert MapTop > 0
        If Rnd < (MapTop - 1) / Size Then
            Exterminate Map(K)
        Else
            'Copy
            If Rnd < CopyPoss Then
                AutoLoad Map(K)
            End If
        End If
    End With
    K = K + 1
Loop
Me.Caption = MapTop
End Sub

Private Sub LoadDaniel(ID As Integer, ParentID As Integer)
'Debug.Assert ID <> 0
Debug.Assert Map(MapTop) = 0
With Daniel(ID)
    .Exist = True
    .X = Daniel(ParentID).X
    .Y = Daniel(ParentID).Y
    .Rad = Daniel(ParentID).Rad
    .F = Daniel(ParentID).F
    .R = Daniel(ParentID).R
    .G = Daniel(ParentID).G
    .B = Daniel(ParentID).B
    Debug.Assert .F <> 0
    Dim Bounce As Boolean
    Bounce = True
    Do While Bounce
        Bounce = False
        Select Case Rnd * 6
            Case 0 To 1
                .R = .R + 1
                If .R = 16 Then
                    .R = 15
                    Bounce = True
                End If
            Case 1 To 2
                .G = .G + 1
                If .G = 16 Then
                    .G = 15
                    Bounce = True
                End If
            Case 2 To 3
                .B = .B + 1
                If .B = 16 Then
                    .B = 15
                    Bounce = True
                End If
            Case 3 To 4
                .R = .R - 1
                If .R = 0 Then
                    .R = 1
                    Bounce = True
                End If
            Case 4 To 5
                .G = .G - 1
                If .G = 0 Then
                    .G = 1
                    Bounce = True
                End If
            Case Else
                .B = .B - 1
                If .B = 0 Then
                    .B = 1
                    Bounce = True
                End If
        End Select
    Loop
End With
Map(MapTop) = ID
MapTop = MapTop + 1
'Debug.Print "+" & ID
Debug.Assert Map(MapTop) = 0
End Sub

Private Sub Exterminate(ID As Integer)
If MapTop = 1 Then
    Form_Load
    Cls
    Exit Sub
End If
'Debug.Assert ID <> 0
'Debug.Print "Exterminate " & ID
Dim K As Integer
With Daniel(ID)
    .Exist = False
    .X = 0
    .Y = 0
    .Rad = 0
    .F = 0
    .R = 0
    .G = 0
    .B = 0
End With
Dim Pos As Integer
Pos = 0
Do Until Map(Pos) = ID
    Pos = Pos + 1
Loop
For K = Pos To MapTop - 1
    Map(K) = Map(K + 1)
Next K
MapTop = MapTop - 1
End Sub

Private Sub AutoLoad(ParentID As Integer)
Dim K As Integer
Do While Daniel(K).Exist
    K = K + 1
Loop
LoadDaniel K, ParentID
End Sub

Private Sub ShowMap()
Debug.Print "MapTop=" & MapTop
Dim K As Integer
For K = 0 To MapTop - 1
    Debug.Print Map(K);
Next K
Debug.Print
End Sub

Private Sub ShowDaniel(ID)
With Daniel(ID)
    Debug.Print .Exist
    Debug.Print .X
    Debug.Print .Y
    Debug.Print .Rad
    Debug.Print .F
End With
End Sub


Attribute VB_Name = "Hall"
Option Explicit

Private Type typPlayer
    Name As String
    Colour As String
End Type

Public Player(1 To 99) As typPlayer

Public Function RoomInfo(PlayerNum As Integer) As String
Dim k As Integer
RoomInfo = PlayerNum
For k = 1 To PlayerNum
    RoomInfo = RoomInfo & " " & Player(k).Name & " " & NumColour(Player(k).Colour)
Next k
End Function

Public Sub Rename(ID As Integer, NewName As String)
Player(ID).Name = NewName
End Sub

Private Function NumColour(RgbColour As String) As String
Dim S() As String
S = Split(RgbColour, ",")
NumColour = RGB(S(0), S(1), S(2))
End Function

Public Sub Recolour(ID As Integer, NewColour As String)
Player(ID).Colour = NewColour
End Sub

Public Sub AddPlayer(PlayerNum As Integer)
With Player(PlayerNum)
    .Name = PlayerNum
    .Colour = "255,255,255"
End With
End Sub

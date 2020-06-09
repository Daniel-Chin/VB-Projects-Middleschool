VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Daniels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public X As Long
Public Y As Long
Public Rad As Double

Private Sub Class_Terminate()
X = 0
Y = 0
Rad = 0
End Sub

Public Sub Birth(Parent As Daniels)
With Parent
    X = .X
    Y = .Y
    Rad = .Rad
End With
End Sub


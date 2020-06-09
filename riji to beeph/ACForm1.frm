VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   1488
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   1920
   LinkTopic       =   "Form1"
   ScaleHeight     =   1488
   ScaleWidth      =   1920
   Visible         =   0   'False
   Begin VB.Timer T 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Old As Integer

Private Sub T_Timer()
Dim SY As Integer
SY = GetSetting("Jack", "RiJi", "Access")
If SY = 0 Then
    End
Else
    If SY = Old - 1 Then
        Unload Me
        End
    Else
        SY = SY - 1
        SaveSetting "Jack", "RiJi", "Access", SY
    End If
End If
Old = SY
End Sub

VERSION 5.00
Begin VB.Form Heartbeat_ 
   Caption         =   "Heartbeat"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   3744
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Heartbeat 
      Interval        =   714
      Left            =   1440
      Top             =   1080
   End
End
Attribute VB_Name = "Heartbeat_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ILoveGabrielle = True

Private Sub Heartbeat_Timer()
Debug.Assert ILoveGabrielle
End Sub

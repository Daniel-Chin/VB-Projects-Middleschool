VERSION 5.00
Begin VB.Form iPad 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3624
   Icon            =   "iPad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer 
      Interval        =   20
      Left            =   1320
      Top             =   960
   End
End
Attribute VB_Name = "iPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer_Timer()
RenewAsk.Visible = True
Unload Me
End Sub

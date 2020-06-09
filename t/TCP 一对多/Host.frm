VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2964
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5316
   LinkTopic       =   "Form1"
   ScaleHeight     =   2964
   ScaleWidth      =   5316
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   1920
      Top             =   960
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
      LocalPort       =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Winsock1(0).Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
For k = 0 To Winsock1.UBound
    Winsock1(k).Close
Next k
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Load Winsock1(Winsock1.UBound + 1)
Winsock1(Winsock1.UBound).Accept requestID
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
MsgBox Index
End Sub

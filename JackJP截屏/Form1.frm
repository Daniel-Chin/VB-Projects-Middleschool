VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   3744
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Sub Form_Load()
    Randomize
    Me.AutoRedraw = True
    Dim w As Long, h As Long
    Dim hwndSrc As Long
    Dim hSrcDC As Long
    
    hwndSrc = GetDesktopWindow()
    hSrcDC = GetWindowDC(hwndSrc)
    w = Screen.Width \ Screen.TwipsPerPixelX
    h = Screen.Height \ Screen.TwipsPerPixelY
    
    Call BitBlt(hdc, 0, 0, w, h, hSrcDC, 0, 0, SRCCOPY)
    Call ReleaseDC(hwndSrc, hSrcDC)
    
    SavePicture Me.Image, "D:\JackJP\" & Rnd & ".jpg"
End Sub

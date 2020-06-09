Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Const SM_CXSCREEN = 0
Const SM_CYSCREEN = 1

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long

Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000

Private Type DEVMODE
dmDeviceName As String * CCDEVICENAME
dmSpecVersion As Integer
dmDriverVersion As Integer
dmSize As Integer
dmDriverExtra As Integer
dmFields As Long
dmOrientation As Integer
dmPaperSize As Integer
dmPaperLength As Integer
dmPaperWidth As Integer
dmScale As Integer
dmCopies As Integer
dmDefaultSource As Integer
dmPrintQuality As Integer
dmColor As Integer
dmDuplex As Integer
dmYResolution As Integer
dmTTOption As Integer
dmCollate As Integer
dmFormName As String * CCFORMNAME
dmUnusedPadding As Integer
dmBitsPerPel As Integer
dmPelsWidth As Long
dmPelsHeight As Long
dmDisplayFlags As Long
dmDisplayFrequency As Long
End Type
Dim DevM As DEVMODE


Sub ChangeRes(iWidth As Single, iHeight As Single)
Dim a As Boolean
 Dim i As Integer
 Dim b As Long
i = 0
Do
a = EnumDisplaySettings(0&, i, DevM)
 i = i + 1
Loop Until (a = False)
 DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
 DevM.dmPelsWidth = iWidth
 DevM.dmPelsHeight = iHeight
 ChangeDisplaySettings DevM, 0
End Sub

Private Sub Main()
Call ChangeRes(Val(InputBox("x", "", "640")), Val(InputBox("x", "", "480")))
End Sub

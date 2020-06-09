Attribute VB_Name = "MD"
Option Explicit

Public hHook As Long

Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public LastwParam As Long, LastlParam As Long

Public Function HookFunc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If ncode >= 0 Then
    HookFunc = 0
    LastwParam = wParam
    LastlParam = lParam
        HookFunc = 1
End If
Call CallNextHookEx(hHook, ncode, wParam, lParam)
End Function

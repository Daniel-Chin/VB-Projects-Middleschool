VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private IsHooking As Boolean

Public Function Hook() As Boolean
MD.hHook = MD.SetWindowsHookEx(2, AddressOf HookFunc, MD.GetModuleHandle(App.Path & "\" & App.EXEName), 0)
Hook = MD.hHook <> 0
IsHooking = True
End Function

Public Sub Unhook()
MD.UnhookWindowsHookEx MD.hHook
IsHooking = False
End Sub

Private Sub Class_Terminate()
If IsHooking Then
     Unhook
End If
End Sub

Public Sub Listen(ByRef wParam As Long, ByRef lParam As Long)
Do While wParam = 0
    GetData wParam, lParam
    Sleep 10
    DoEvents
Loop
End Sub

Public Sub GetData(ByRef wParam As Long, ByRef lParam As Long)
If MD.LastwParam <> 0 Then
    wParam = MD.LastwParam
    lParam = MD.LastlParam
    MD.LastwParam = 0
    MD.LastlParam = 0
End If
End Sub

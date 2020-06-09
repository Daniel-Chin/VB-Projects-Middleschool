Attribute VB_Name = "Bas"
Private Declare Function SetWindowsHookEx _
 Lib "user32" Alias "SetWindowsHookExA" ( _
 ByVal idHook As Long, _
 ByVal lpfn As Long, _
 ByVal hmod As Long, _
 ByVal dwThreadId As Long) As Long
  
Private Declare Function CallNextHookEx Lib "user32" ( _
 ByVal hHook As Long, _
 ByVal nCode As Long, _
 ByVal wParam As Long, _
 ByVal lParam As Long) As Long
 
Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
 ByVal hHook As Long) As Long
 
Private Declare Sub CopyMemory _
 Lib "kernel32" Alias "RtlMoveMemory" ( _
 pDest As Any, _
 pSource As Any, _
 ByVal cb As Long)
 
Private Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type
 
Private Const HC_ACTION = 0&
Private Const WH_KEYBOARD_LL = 13&
Private Const VK_LWIN = &H5B&
Private Const VK_RWIN = &H5C&
 
Private hKeyb As Long
 
Private Function KeybCallback(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static udtHook As KBDLLHOOKSTRUCT
    
    If (Code = HC_ACTION) Then
        'Copy the keyboard data out of the lParam (which is a pointer)
        Call CopyMemory(udtHook, ByVal lParam, Len(udtHook))
        Select Case udtHook.vkCode
            Case VK_LWIN, VK_RWIN
                KeybCallback = 1
                Exit Function
        End Select
    End If
    KeybCallback = CallNextHookEx(hKeyb, Code, wParam, lParam)
End Function
 
Public Sub HookKeyboard()
    UnhookKeyboard
    hKeyb = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeybCallback, App.hInstance, 0&)
End Sub
 
Public Sub UnhookKeyboard()
    If hKeyb <> 0 Then
        Call UnhookWindowsHookEx(hKeyb)
        hKeyb = 0
    End If
End Sub

Attribute VB_Name = "ModWndProc"
 '*************************************************************************
 '**说    明：紫水晶工作室 http://www.m5home.com/
 '**创 建 人：马大哈
'**日    期：2005年04月13日
'**描    述：一个使用子类技术得到鼠标滚轮状态的简单例子
'*************************************************************************
Option Explicit

 Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 Public Const GWL_WNDPROC = (-4)
 Public Const WM_GETTEXT = &HD
 Public Const WM_MOUSEWHEEL = &H20A

 Public PrevWndProc As Long
 Public MouseW As Long

 Public Function SubWndProc(ByVal Hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

 Select Case MSG         '在这里进行过滤.如果知道其他的消息,也可以在这里过滤.

     Case WM_MOUSEWHEEL
         
         If wParam > 0 Then
             MouseW = 1
         ElseIf wParam < 0 Then
             MouseW = -1
         End If
         
 End Select

 SubWndProc = CallWindowProc(PrevWndProc, Hwnd, MSG, wParam, lParam)     '其它消息不管

End Function
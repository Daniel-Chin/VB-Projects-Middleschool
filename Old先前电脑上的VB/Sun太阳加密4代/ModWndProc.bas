Attribute VB_Name = "ModWndProc"
 '*************************************************************************
 '**˵    ������ˮ�������� http://www.m5home.com/
 '**�� �� �ˣ�����
'**��    �ڣ�2005��04��13��
'**��    ����һ��ʹ�����༼���õ�������״̬�ļ�����
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

 Select Case MSG         '��������й���.���֪����������Ϣ,Ҳ�������������.

     Case WM_MOUSEWHEEL
         
         If wParam > 0 Then
             MouseW = 1
         ElseIf wParam < 0 Then
             MouseW = -1
         End If
         
 End Select

 SubWndProc = CallWindowProc(PrevWndProc, Hwnd, MSG, wParam, lParam)     '������Ϣ����

End Function
Attribute VB_Name = "modMain"
'====================================================
'描述:      提供一些全局通用的函数，如窗口消息处理等
'作者:      冰棍
'文件:      modMain.bas
'====================================================

Option Explicit

'调用系统的消息处理过程
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
    ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'获取系统参数信息
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    
'描述:      获取本程序的路径。如果路径后面缺少"\"，则自动加上
'返回值:    以"\"结尾的路径
Public Function GetAppPath() As String
    GetAppPath = App.Path
    If Right(GetAppPath, 1) <> "\" Then
        GetAppPath = GetAppPath & "\"
    End If
End Function

'描述:      修复主窗口最大化全屏的问题
'参数:      hWnd: 窗口句柄
'.          uMsg: 消息值
'.          wParam, lParam: 消息的参数
'返回值:    消息处理返回值
Public Function MainWindowMaximizeFixProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_GETMINMAXINFO Then                                                         '窗口尝试获取最大、最小化信息
        Dim mmi             As MINMAXINFO                                                       '最大、最小化信息
        Dim rectWorkArea    As RECT                                                             '屏幕工作区大小
        
        'lParam为指向MINMAXINFO的指针
        CopyMemory mmi, ByVal lParam, ByVal Len(mmi)
        SystemParametersInfo SPI_GETWORKAREA, ByVal 0, rectWorkArea, ByVal 0                    '获取屏幕工作区大小
        mmi.ptMaxSize.Y = rectWorkArea.bottom - rectWorkArea.Top
        CopyMemory ByVal lParam, mmi, ByVal Len(mmi)                                            '更改最大化信息中的大小信息，修复窗口最大化的时候会全屏的问题
        
        MainWindowMaximizeFixProc = 0                                                           '处理这个消息之后需要返回0
        Exit Function
    End If
    MainWindowMaximizeFixProc = CallWindowProc(GetPropA(hWnd, "PrevWndProc"), hWnd, uMsg, wParam, lParam)
End Function

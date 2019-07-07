Attribute VB_Name = "modMain"
'====================================================
'描述:      提供一些全局通用的函数，如窗口消息处理等
'作者:      冰棍
'文件:      modMain.bas
'====================================================

Option Explicit

'调用系统的消息处理过程
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'获取系统参数信息
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public DebugProgramInfo     As PROCESS_INFORMATION                                      '正在调试中的进程信息
    
'描述:      获取本程序的路径。如果路径后面缺少"\"，则自动加上
'返回值:    以"\"结尾的路径
Public Function GetAppPath() As String
    GetAppPath = App.Path
    If Right(GetAppPath, 1) <> "\" Then
        GetAppPath = GetAppPath & "\"
    End If
End Function

'描述:      判断进程中是否存在有指定PID的进程
'参数:      hProcess: 进程句柄
'返回值:    指定的进程是否存在
Public Function ProcessExists(ByVal hProcess As Long) As Boolean
    Dim ret         As Long
    
    ret = WaitForSingleObject(hProcess, 0)                                                  '判断进程是否退出
    ProcessExists = (ret = WAIT_TIMEOUT)                                                    '当返回值为超时说明进程仍在运行
End Function

'描述:      在使用MsgBox前先关闭皮肤，否则会让MsgBox很难看 :)
'参数:      MsgBox的前三个参数
'返回值:    MsgBox的返回值
Public Function NoSkinMsgBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String) As VbMsgBoxResult
    frmMain.SkinFramework.AutoApplyNewThreads = False
    frmMain.SkinFramework.AutoApplyNewWindows = False
    NoSkinMsgBox = MsgBox(Prompt, Buttons, Title)
    frmMain.SkinFramework.AutoApplyNewThreads = True
    frmMain.SkinFramework.AutoApplyNewWindows = True
End Function

'描述:      把两个16位的数联合成一个32位的Long型数
'参数:      wLow, wHigh: 分别是低16位和高16位
'返回值:    合成的数
Public Function MakeLong(wLow As Long, wHigh As Long) As Long
    MakeLong = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

'描述:      获取一个32位数的高16位
'参数:      lValue: 数值
'返回值:    高16位的数值
Public Function HiWord(lValue As Long) As Integer
    If lValue And &H80000000 Then
        HiWord = (lValue \ 65535) - 1
    Else
        HiWord = lValue \ 65535
    End If
End Function

'描述:      获取一个32位数的低16位
'参数:      lValue: 数值
'返回值:    低16位的数值
Public Function LoWord(lValue As Long) As Integer
    If lValue And &H8000& Then
        LoWord = &H8000 Or (lValue And &H7FFF&)
    Else
        LoWord = lValue And &HFFFF&
    End If
End Function

'描述:      通过wParam计算出Shift值
'参数:      wParam: wParam值
'返回值:    Shift值
Public Function GetShiftValue(wParam As Long) As Long
    GetShiftValue = (wParam And MK_CONTROL) Or (wParam And MK_SHIFT)
End Function

'描述:      修复主窗口最大化全屏和在任务栏的右键菜单无法关闭的问题
'参数:      hWnd: 窗口句柄
'.          uMsg: 消息值
'.          wParam, lParam: 消息的参数
'返回值:    消息处理返回值
Public Function MainWindowMaximizeCloseFixProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_GETMINMAXINFO Then                                                         '窗口尝试获取最大、最小化信息
        Dim mmi             As MINMAXINFO                                                       '最大、最小化信息
        Dim rectWorkArea    As RECT                                                             '屏幕工作区大小
        
        'lParam为指向MINMAXINFO的指针
        CopyMemory mmi, ByVal lParam, ByVal Len(mmi)
        SystemParametersInfo SPI_GETWORKAREA, ByVal 0, rectWorkArea, ByVal 0                    '获取屏幕工作区大小
        mmi.ptMaxSize.Y = rectWorkArea.bottom - rectWorkArea.Top
        CopyMemory ByVal lParam, mmi, ByVal Len(mmi)                                            '更改最大化信息中的大小信息，修复窗口最大化的时候会全屏的问题
        
        MainWindowMaximizeCloseFixProc = 0                                                      '处理这个消息之后需要返回0
        Exit Function
    ElseIf uMsg = WM_SYSCOMMAND Then                                                        '在任务栏使用右键菜单关闭
        If wParam = SC_CLOSE Then
            Dim WindowObj   As Object                                                           '对应的窗体物件
            
            CopyMemory ByVal VarPtr(WindowObj), GetPropA(hWnd, "WindowObj"), ByVal 4            '获取该窗口对应的Form
            Unload WindowObj                                                                    '卸载Form
        End If
    End If
    MainWindowMaximizeCloseFixProc = CallWindowProc(GetPropA(hWnd, "PrevWndProc"), hWnd, uMsg, wParam, lParam)
End Function

'描述:      显示“打开”通用对话框
'参数:      hWnd: 调用该函数的窗口句柄
'.          Title: 对话框标题
'.          Filter: 文件过滤器，使用vbNullChar来隔开每个过滤器
'返回值:    如果操作取消或者出错，返回""；否则返回选择的文件路径
Public Function ShowOpen(hWnd As Long, Title As String, Filter As String) As String
    Dim ofn                 As OPENFILENAME                                                 '对话框信息
    
    Filter = Filter & vbNullChar
    With ofn                                                                                '设置对话框信息
        .lStructSize = Len(ofn)
        .hWndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrFilter = Filter
        .lpstrFile = String(MAX_PATH, vbNullChar)                                               '设置文件名缓冲区
        .nMaxFile = MAX_PATH
        .lpstrFileTitle = Space(MAX_PATH - 1)
        .nMaxFileTitle = MAX_PATH
        .lpstrTitle = Title
        .Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
        .nFilterIndex = 0
    End With
    
    frmMain.SkinFramework.AutoApplyNewThreads = False                                       '暂时禁用皮肤控件
    frmMain.SkinFramework.AutoApplyNewWindows = False
    If GetOpenFileNameW(ofn) = 1 Then                                                       '显示保存对话框
        ShowOpen = Split(ofn.lpstrFile, vbNullChar)(0)                                          '在'\0'处截断字符串
    End If
    frmMain.SkinFramework.AutoApplyNewThreads = True                                        '重新启用皮肤控件
    frmMain.SkinFramework.AutoApplyNewWindows = True
End Function

'描述:      显示“保存”通用对话框
'参数:      hWnd: 调用该函数的窗口句柄
'.          DefaultName: 默认的文件名
'.          Title: 对话框标题
'.          Filter: 文件过滤器，使用vbNullChar来隔开每个过滤器
'返回值:    如果操作取消或者出错，返回""；否则返回选择的文件路径
Public Function ShowSave(hWnd As Long, DefaultName As String, Title As String, Filter As String) As String
    Dim ofn                 As OPENFILENAME                                                 '对话框信息
    
    DefaultName = DefaultName & String(MAX_PATH - Len(DefaultName), vbNullChar)             '字符串结尾加上足够数量的'\0'，作为缓冲区
    Filter = Filter & vbNullChar                                                            '字符串末尾必须是'\0'
    With ofn                                                                                '设置对话框信息
        .lStructSize = Len(ofn)
        .hWndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrFilter = Filter
        .lpstrFile = DefaultName
        .nMaxFile = MAX_PATH
        .lpstrFileTitle = Space(MAX_PATH - 1)
        .nMaxFileTitle = MAX_PATH
        .lpstrTitle = Title
        .Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
        .nFilterIndex = 0
    End With
    
    frmMain.SkinFramework.AutoApplyNewThreads = False                                       '暂时禁用皮肤控件
    frmMain.SkinFramework.AutoApplyNewWindows = False
    If GetSaveFileNameW(ofn) = 1 Then                                                       '显示保存对话框
        ShowSave = Split(ofn.lpstrFile, vbNullChar)(0)                                          '在'\0'处截断字符串
    End If
    frmMain.SkinFramework.AutoApplyNewThreads = True                                        '重新启用皮肤控件
    frmMain.SkinFramework.AutoApplyNewWindows = True
End Function

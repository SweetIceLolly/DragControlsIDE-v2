Attribute VB_Name = "modTooltips"
'====================================================
'描述:      工具提示文本模块，创建支持多行的、可以自定义样式的工具提示文本
'作者:      冰棍
'文件:      modTooltips.bas
'====================================================

Option Explicit

Dim hWndTip         As Long                             '工具提示文本窗口句柄

'描述:      创建工具提示文本窗口（在程序初始化时调用）
'返回值:    创建的工具提示文本窗口句柄
Public Function CreateToolTip() As Long
    hWndTip = CreateWindowExA(0, "tooltips_class32", vbNullString, WS_POPUP Or TTS_ALWAYSTIP, _
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0, 0, App.hInstance, 0)
    
    SendMessageA hWndTip, TTM_SETTIPBKCOLOR, ByVal &H454242, 0                  '设置背景颜色
    SendMessageA hWndTip, TTM_SETTIPTEXTCOLOR, ByVal &HF0F0F0, 0                '设置文本颜色
    SendMessageA hWndTip, TTM_SETMAXTIPWIDTH, 0, ByVal &HFFFFFFFF               '设置为多行的
    
    CreateToolTip = hWndTip
End Function

'描述:      关闭工具提示文本窗口（在程序退出时调用）
Public Sub DestroyToolTip()
    DestroyWindow hWndTip
End Sub

'描述:      为指定的控件添加工具提示文本
'参数:      TargetWindow: 需要添加工具提示文本的窗口
'.          Tooltip: 工具提示文本
'.          Title: 可选的，指定工具提示文本的标题
'.          Icon: 可选的，指定工具提示文本的图标
'返回值:    返回1表示添加成功，0表示添加失败
Public Function CtlAddToolTip(TargetWindow As Long, Tooltip As String, _
    Optional Title As String, Optional Icon As Tooltip_Icon = 0) As Long
    
    Dim ti          As TTTOOLINFO
    Dim tmpStr()    As Byte
    
    tmpStr = StrconvEx(Tooltip)
    With ti
        .cbSize = Len(ti)
        .hWnd = TargetWindow
        .uFlags = TTF_IDISHWND Or TTF_SUBCLASS
        .uID = TargetWindow
        .lpszText = VarPtr(tmpStr(0))
    End With
    
    CtlAddToolTip = SendMessageA(hWndTip, TTM_ADDTOOL, 0, ByVal VarPtr(ti))
    
    If Len(Title) > 1 Then
        tmpStr = StrconvEx(Title)
        SendMessageA hWndTip, TTM_SETTITLE, ByVal Icon, ByVal VarPtr(tmpStr(0))
    End If
End Function


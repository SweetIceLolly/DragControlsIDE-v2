Attribute VB_Name = "modGdbParser"
'====================================================
'描述:      提供解析gdb输出的函数
'作者:      冰棍
'文件:      modGdbParser.bas
'====================================================

Option Explicit

'定义调用堆栈信息结构
Public Type CallStackInfoStruct
    Address                 As String                                       '地址
    Args                    As String                                       '参数
    File                    As String                                       '文件
    Line                    As Long                                         '行号 (-1则代表要从文件浏览器显示该文件)
End Type

'定义模块信息结构
Public Type ModuleInfoStruct
    File                    As String                                       '模块文件
    From                    As String                                       '从（地址）
    To                      As String                                       '到（地址）
End Type

'定义线程信息结构
Public Type ThreadInfoStruct
    Id                      As String                                       '线程ID
    Frame                   As String                                       '地址
End Type

'描述:      分析gdb的堆栈输出
'参数:      strCallStack: 需要分析的调用堆栈输出
'返回值:    存储着调用堆栈信息的结构
Public Function ParseCallStackString(strCallStack As String) As CallStackInfoStruct
    'On Error Resume Next
    Dim StrPos              As Long                                         '查找到的字符串的位置
    Dim Info                As CallStackInfoStruct
    
    If Mid(strCallStack, Len(strCallStack)) = vbCr Then                                 '去掉字符串结尾的换行符
        strCallStack = Left(strCallStack, Len(strCallStack) - 1)
    End If
    
    '有准确地址；有对应的动态库函数及文件
    If strCallStack Like "[#]* * in *(*)* from *:[\/]*" Then
        '例子: #2  0x76926359 in KERNEL32!BaseThreadInitThunk () from C:\WINDOWS\SysWOW64\kernel32.dll
        Info.Line = -1                                                                      '标记为要从文件浏览器显示该文件
        StrPos = InStrRev(strCallStack, ":/")                                               '查找字符串中的“:/"，以分割出文件名
        If StrPos = 0 Then                                                                  '找不到“:/”（新版gdb）就尝试查找“:\”
            StrPos = InStrRev(strCallStack, ":\")
        End If
        Info.File = Right(strCallStack, Len(strCallStack) - _
            InStrRev(strCallStack, " from ", StrPos) - 5)                                   '（#n  addr in func (args) from [file]）
        strCallStack = Left(strCallStack, Len(strCallStack) - Len(Info.File) - 6)           '（[#n  addr in func (args)] from file）
        StrPos = InStrRev(strCallStack, " (")                                               '查找字符串中的“ (”，以分割出参数
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '（#n  addr in func ([args])）
        StrPos = InStr(strCallStack, " 0x")                                                 '查找字符串中的“ 0x”，以分割出地址
        Info.Address = Mid(strCallStack, StrPos + 1, _
            Len(strCallStack) - StrPos - Len(Info.Args) - 3)                                '（#n  [addr in func] (args)）
            
    '---------------------------------------------------------
    '有准确地址；有对应文件
    ElseIf strCallStack Like "[#]* * in *(*)* at *:[\/]*" Then
        '例子: #1  0x0040144c in main () at C:\(aa) bb\.cpp:6
        StrPos = InStrRev(strCallStack, ":")                                                '查找字符串中的“:”，以分割出行号
        Info.Line = CLng(Right(strCallStack, Len(strCallStack) - StrPos))                   '（#n  addr in func (args) at file:[line]）
        strCallStack = Left(strCallStack, StrPos - 1)                                       '（[#n  addr in func (args) at file]:line）
        StrPos = InStrRev(strCallStack, ":/")                                               '查找字符串中的“:/"，以分割出文件名
        If StrPos = 0 Then                                                                  '找不到“:/”（新版gdb）就尝试查找“:\”
            StrPos = InStrRev(strCallStack, ":\")
        End If
        Info.File = Right(strCallStack, Len(strCallStack) - _
            InStrRev(strCallStack, " at ", StrPos) - 3)                                     '（#n  addr in func (args) at [file]）
        strCallStack = Left(strCallStack, Len(strCallStack) - Len(Info.File) - 4)           '（[#n  addr in func (args)] at file）
        StrPos = InStrRev(strCallStack, " (")                                               '查找字符串中的“ (”，以分割出参数
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '（#n  addr in func ([args])）
        StrPos = InStr(strCallStack, " 0x")                                                 '查找字符串中的“ 0x”，以分割出地址
        Info.Address = Mid(strCallStack, StrPos + 1, _
            Len(strCallStack) - StrPos - Len(Info.Args) - 3)                                '（#n  [addr in func] (args)）
            
    '---------------------------------------------------------
    '有准确地址；无对应文件
    ElseIf strCallStack Like "[#]* * in *(*)*" Then
        '例子: #1  0x00403c44 in main ()
        StrPos = InStrRev(strCallStack, " (")                                               '查找字符串中的“ (”，以分割出参数
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '（#n  addr in func ([args])）
        StrPos = InStr(strCallStack, " 0x")                                                 '查找字符串中的“ 0x”，以分割出地址
        Info.Address = Mid(strCallStack, StrPos + 1, _
            Len(strCallStack) - StrPos - Len(Info.Args) - 3)                                '（#n  [addr in func] (args)）
    
    '---------------------------------------------------------
    '无准确地址；有对应文件
    ElseIf strCallStack Like "[#]* * (*)* at *:[\/]*" Then
        '例子: #0  aaa (a=1, b=2, c=3, d=4) at C:\(aa) bb\.cpp:6
        StrPos = InStrRev(strCallStack, ":")                                                '查找字符串中的“:”，以分割出行号
        Info.Line = CLng(Right(strCallStack, Len(strCallStack) - StrPos))                   '（#n  func (args) at file:[line]）
        strCallStack = Left(strCallStack, StrPos - 1)                                       '（[#n  func (args) at file]:line）
        StrPos = InStrRev(strCallStack, ":/")                                               '查找字符串中的“:/"，以分割出文件名
        If StrPos = 0 Then                                                                  '找不到“:/”（新版gdb）就尝试查找“:\”
            StrPos = InStrRev(strCallStack, ":\")
        End If
        Info.File = Right(strCallStack, Len(strCallStack) - _
            InStrRev(strCallStack, " at ", StrPos) - 3)                                     '（#n  func (args) at [file]）
        strCallStack = Left(strCallStack, Len(strCallStack) - Len(Info.File) - 4)           '（[#n  func (args)] at file）
        StrPos = InStr(strCallStack, " ")                                                   '查找字符串中的“ ”，以去掉开头的序号
        strCallStack = Trim(Right(strCallStack, Len(strCallStack) - StrPos))                '（#n  [func (args)]）
        StrPos = InStrRev(strCallStack, " (")                                               '查找字符串中的“ (”，以分割出参数
        Info.Address = Left(strCallStack, StrPos - 1)                                       '（[func] (args)）
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '（func ([args])）
    
    '---------------------------------------------------------
    '无准确地址；无对应文件
    ElseIf strCallStack Like "[#]* * (*)*" Then
        '例子: #1  aaa (a=1, b=2, c=3, d=4)
        StrPos = InStr(strCallStack, " ")                                                   '查找字符串中的“ ”，以去掉开头的序号
        strCallStack = Trim(Right(strCallStack, Len(strCallStack) - StrPos))                '（#n  [func (args)]）
        StrPos = InStrRev(strCallStack, " (")                                               '查找字符串中的“ (”，以分割出参数
        Info.Address = Left(strCallStack, StrPos - 1)                                       '（[func] (args)）
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '（func ([args])）
    
    '---------------------------------------------------------
    '遭遇C++巨佬！放弃解析，直接添加到列表里
    Else
        StrPos = InStr(strCallStack, " ")                                                   '查找字符串中的“ ”，以去掉开头的序号
        Info.Address = Trim(Right(strCallStack, Len(strCallStack) - StrPos))                '（#n  [func (args)]）
    End If
    
    Info.File = Replace(Info.File, "/", "\")                                            '把地址里的“/”替换成“\”
    ParseCallStackString = Info
End Function

'描述:      分析gdb的模块输出
'参数:      strModule: 需要分析的模块输出
'返回值:    存储着模块信息的结构
Public Function ParseModuleString(strModule As String) As ModuleInfoStruct
    'on error resume next
    
    '例子：
    '(gdb) info sharedlibrary
    'From        To          Syms Read   Shared Object Library
    '0x76920000  0x769e47b0  Yes (*)     C:\WINDOWS\SysWOW64\kernel32.dll
    '0x77051000  0x7724bfd0  Yes (*)     C:\WINDOWS\SysWOW64\KernelBase.dll
    '0x766c1000  0x7677e764  Yes (*)     C:\WINDOWS\SysWOW64\msvcrt.dll
    '(*): Shared library is missing debugging information.
    '(gdb)

    Dim StrPos              As Long                                     '查找到的字符串位置
    Dim Info                As ModuleInfoStruct
    
    If Mid(strModule, Len(strModule)) = vbCr Then                       '去掉字符串结尾的换行符
        strModule = Left(strModule, Len(strModule) - 1)
    End If
    
    '检测字符串是否符合格式
    If strModule Like "*0x* 0x* * C:[\/]*" Then
        StrPos = InStr(strModule, "0x")                                     '搜索第一个“0x”，获取“从”地址
        Info.From = Mid(strModule, StrPos, 10)
        StrPos = InStr(StrPos + 10, strModule, "0x")                        '搜索第二个“0x”，获取“到”地址
        Info.To = Mid(strModule, StrPos, 10)
        StrPos = InStrRev(strModule, ":\")                                  '从结尾向前搜索“:\”或“:/”，
        If StrPos = 0 Then
            StrPos = InStrRev(strModule, ":/")
        End If
        StrPos = InStrRev(strModule, " ", StrPos)                           '从找到的位置向前查找空格，获取路径
        Info.File = Mid(strModule, StrPos + 1, Len(strModule) - StrPos)
    End If
    
    Info.File = Replace(Info.File, "/", "\")                            '把地址里的“/”替换成“\”
    ParseModuleString = Info
End Function

Public Function ParseThreadString(strThread As String) As ThreadInfoStruct
    'on error resume next
    
    '例子：
    '(gdb) info threads
    '  Id   Target Id         Frame
    '  2    Thread 19152.0x17a0 0x77af3a4c in ?? ()
    '* 1    Thread 19152.0x4794 main () at C:\Users\12574\Documents\MyProjects\te\te.cpp:2
    '(gdb)
    
    Dim StrPos              As Long                                     '查找到的字符串位置
    Dim StrPos2             As Long
    Dim Info                As ThreadInfoStruct
    
    If Mid(strThread, Len(strThread)) = vbCr Then                       '去掉字符串结尾的换行符
        strThread = Left(strThread, Len(strThread) - 1)
    End If
    
    '检测字符串是否符合格式
    If strThread Like "* * *.0x* *" Then
        StrPos = InStr(strThread, ".0x")
        StrPos = InStr(StrPos, strThread, " ")                              '从“.0x”向后搜索“ ”
        Info.Frame = Right(strThread, Len(strThread) - StrPos)              '截取空格后的内容作为地址
        StrPos2 = InStrRev(strThread, " ", StrPos - 1) + 1                  '从“.0x”向前搜索“ ”
        Info.Id = Mid(strThread, StrPos2, StrPos - StrPos2)                 '截取“.0x”前面的空格和后面的空格中间的文本作为ID
    End If
    
    ParseThreadString = Info
End Function

Attribute VB_Name = "modGdbParser"
Option Explicit

'定义调用堆栈信息结构
Public Type CallStackInfoStruct
    Address                                     As String                                   '地址
    Args                                        As String                                   '参数
    File                                        As String                                   '文件
    Line                                        As Long                                     '行号 (-1则代表要从文件浏览器显示该文件)
End Type

'描述:      分析gdb的堆栈输出
'参数:      strCallStack: 需要分析的调用堆栈输出
'返回值:    存储着调用堆栈信息的结构
Public Function ParseCallStackString(strCallStack As String) As CallStackInfoStruct
    'On Error Resume Next
    
    Dim StrPos              As Long                                         '查找到的字符串的位置
    Dim BracketLevel        As Long                                         '括号匹配计数，一开始是0，遇到“(”加1, 遇到“)”减1
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



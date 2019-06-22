VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================================
'描述:      管道类，用来与DOS程序进行通讯
'作者:      冰棍
'文件:      clsPipe.cls
'====================================================

Option Explicit

'进程启动信息
Private Type STARTUPINFO
    cb                      As Long
    lpReserved              As Long
    lpDesktop               As Long
    lpTitle                 As Long
    dwX                     As Long
    dwY                     As Long
    dwXSize                 As Long
    dwYSize                 As Long
    dwXCountChars           As Long
    dwYCountChars           As Long
    dwFillAttribute         As Long
    dwFlags                 As Long
    wShowWindow             As Integer
    cbReserved2             As Integer
    lpReserved2             As Long
    hStdInput               As Long
    hStdOutput              As Long
    hStdError               As Long
End Type

'进程信息
Private Type PROCESS_INFORMATION
    hProcess                As Long
    hThread                 As Long
    dwProcessId             As Long
    dwThreadID              As Long
End Type

'安全属性
Private Type SECURITY_ATTRIBUTES
    nLength                 As Long
    lpSecurityDescriptor    As Long
    bInheritHandle          As Long
End Type

'复制句柄
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, _
    lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
'关闭句柄
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'创建管道
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
'获取管道内容而不清空管道里的内容
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, ByVal lpBuffer As Long, ByVal nBufferSize As Long, _
    ByRef lpBytesRead As Long, ByRef lpTotalBytesAvail As Long, ByRef lpBytesLeftThisMessage As Long) As Long

'创建进程
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'获取当前进程的句柄
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
'杀掉进程
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

'等待一个物件执行完毕
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

'往指定文件写入内容
Private Declare Function WriteFile Lib "kernel32" (ByVal hfile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
'往指定文件读取内容
Private Declare Function ReadFile Lib "kernel32" (ByVal hfile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
    ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

'睡觉觉
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'CreateProcess, dwCreationFlags
Private Const NORMAL_PRIORITY_CLASS = &H20&                                     '普通优先级

'STARTUPINFO, dwFlags
Private Const STARTF_USESTDHANDLES = &H100&                                     '使用标准输入输出(stdio)句柄
Private Const STARTF_USESHOWWINDOW = &H1                                        '指定目标程序的窗口是否显示

'DuplicateHandle, dwOptions
Private Const DUPLICATE_SAME_ACCESS = &H2                                       '使复制后的句柄具有同样的权限

'WaitForSingleObject(), 返回值
Private Const WAIT_TIMEOUT = &H102                                              '等待超时，即表示进程仍在运行

Private hInputRead      As Long, hInputWrite    As Long, hInput     As Long     '输入句柄
Private hOutputRead     As Long, hOutputWrite   As Long, hOutput    As Long     '输出句柄

Dim pi                  As PROCESS_INFORMATION                                  '进程信息
Dim bStopOutput         As Boolean                                              '停止获取输出的信号

'接口
Public Event Output(strOutput As String)                                        '目标程序仍在输出
Public Event OutputFinished()                                                   '目标程序输出完成

'描述:      获取当前管道的目标进程的进程句柄
'返回值:    目标进程的进程句柄
Public Property Get hProcess() As Long
    hProcess = pi.hProcess
End Property

'描述:      获取当前管道的目标进程的进程ID
'返回值:    目标进程的进程ID
Public Property Get dwProcessId() As Long
    dwProcessId = pi.dwProcessId
End Property

'描述:      创建Dos输入、输出管道
'参数:      CommandLine: 需要执行的程序
'返回值:    1代表创建成功，0代表创建失败
Public Function InitDosIO(CommandLine As String) As Long
    Dim sa      As SECURITY_ATTRIBUTES                              '安全属性
    Dim si      As STARTUPINFO                                      '进程启动信息
    Dim ret     As Long                                             '函数返回值
    
    With sa                                                         '设置安全属性
        .nLength = Len(sa)
        .bInheritHandle = 1                                             '句柄可继承
        .lpSecurityDescriptor = 0
    End With
    
    '--------------------------------------------------
    '创建输入、输出句柄
    ret = CreatePipe(hInputRead, hInputWrite, sa, 1024)             '输入句柄
    If ret = 0 Then
        InitDosIO = 0
        Exit Function
    End If
    
    ret = CreatePipe(hOutputRead, hOutputWrite, sa, 65536)          '输出句柄
    If ret = 0 Then                                                 '创建失败则关掉刚刚创建的输入句柄
        CloseHandle hInputRead
        CloseHandle hInputWrite
        InitDosIO = 0
        Exit Function
    End If
    
    '转换输入句柄
    ret = DuplicateHandle(GetCurrentProcess(), hInputWrite, GetCurrentProcess(), hInput, 0, 1, DUPLICATE_SAME_ACCESS)
    If ret = 0 Then                                                 '转换失败则关掉刚刚创建的输入、输出句柄
        CloseHandle hInputRead
        CloseHandle hInputWrite
        CloseHandle hOutputRead
        CloseHandle hOutputWrite
        InitDosIO = 0
        Exit Function
    End If
    CloseHandle hInputWrite
    
    '转换输出句柄
    ret = DuplicateHandle(GetCurrentProcess(), hOutputRead, GetCurrentProcess(), hOutput, 0, 1, DUPLICATE_SAME_ACCESS)
    If ret = 0 Then                                                 '转换失败则关掉刚刚创建的输入、输出句柄
        CloseHandle hInputRead
        CloseHandle hInputWrite
        CloseHandle hOutputRead
        CloseHandle hOutputWrite
        InitDosIO = 0
        Exit Function
    End If
    CloseHandle hOutputRead
    
    '--------------------------------------------------
    '创建目标进程
    With si                                                         '设置进程启动信息
        .cb = Len(si)
        .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW         '要求进程使用stdio句柄，并不显示窗体
        .hStdOutput = hOutputWrite                                      '重定向进程输入、输出
        .hStdError = hOutputWrite
        .hStdInput = hInputRead
    End With
    
    If CreateProcess(ByVal 0, CommandLine, sa, sa, ByVal 1, NORMAL_PRIORITY_CLASS, ByVal 0, ByVal 0, si, pi) <> 1 Then
        InitDosIO = 0
    End If
    
    InitDosIO = 1
End Function

'描述:      关闭Dos输入、输出管道
'返回值:    1代表关闭管道成功，0代表关闭管道失败
Public Function CloseDosIO() As Long
    CloseHandle hInputRead
    CloseHandle hInputWrite
    CloseHandle hOutputRead
    CloseHandle hOutputWrite
    TerminateProcess pi.hProcess, 0
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
End Function

'描述:      对目标进程进行输入
'参数:      strInput: 需要输入的内容
'返回值:    1代表输入成功，0代表输入失败
Public Function DosInput(ByVal strInput As String) As Long
    Dim bWritten    As Long                                                 '成功写入的字节数
    Dim InputData() As Byte                                                 '写入的内容
    Dim szInputData As Long                                                 '写入的内容大小
    
    InputData = StrConv(strInput, vbFromUnicode)                            '输入内容转换成字节数组
    szInputData = UBound(InputData) + 1                                     '计算写入内容的大小
    WriteFile hInput, InputData(0), ByVal szInputData, bWritten, 0          '往输入管道写入内容
    
    If bWritten = 0 Then                                                    '居然啥都没写入？？？
        DosInput = 0
        Exit Function
    End If
    DosInput = 1
End Function

'描述:      判断指定的进程是否存在
'返回值:    True: 指定的进程仍然存在, False: 指定的进程未在运行
Private Function IsProcessExists() As Boolean
    Dim ret         As Long
    
    ret = WaitForSingleObject(pi.hProcess, 0)                               '判断进程是否退出
    IsProcessExists = (ret = WAIT_TIMEOUT)                                  '当返回值为超时说明进程仍在运行
End Function

'描述:      获取目标进程的输出
'参数:      lpStrOutput: 存放输出的字符串
'           EndingStr:
'返回值:    1代表获取成功，0代表获取失败
Public Function DosOutput(ByRef strOutput As String, ByVal EndingStr As String) As Long
    Dim ret         As Long                                                 '函数的返回值
    Dim bRead       As Long                                                 '成功读取的内容大小
    Dim OutData()   As Byte                                                 '读取的内容
    Dim szOutData   As Long                                                 '总共需要读取的内容大小
    
    bStopOutput = False
    strOutput = ""
    Do
        If bStopOutput Then                                                     '接收到停止获取输出的信号，函数返回
            RaiseEvent OutputFinished
            DosOutput = 0
            Exit Function
        End If
        If Not IsProcessExists() Then                                           '目标进程已停止运行
            Exit Do
        End If
        ret = PeekNamedPipe(hOutput, 0, 0, ByVal 0, szOutData, ByVal 0)         '查询输出管道里的信息量
        If ret = 0 Then
            DosOutput = 0
            Exit Function
        End If
        
        If szOutData > 0 Then                                                   '如果管道里有可供读取的内容
            ReDim OutData(szOutData)                                                '分配缓冲区接收管道信息
            ret = ReadFile(hOutput, OutData(0), szOutData, bRead, 0)                '从管道里读取内容
            If ret = 0 Then
                DosOutput = 0
                Exit Function
            End If
            
            strOutput = strOutput & StrConv(OutData, vbUnicode)
            strOutput = Left(strOutput, Len(strOutput) - 1)
            strOutput = Replace(strOutput, vbLf, vbCrLf)
        End If
        
        DoEvents                                                                '不要霸占着整条线程
        Sleep 10
        RaiseEvent Output(strOutput)
    Loop Until (Right(strOutput, Len(EndingStr)) = EndingStr)
    RaiseEvent OutputFinished
    DosOutput = 1
End Function

'描述:      中断当前仍在进行中的DosOutput()。余下的输出内容将不会被获取，DosOutput()函数会直接返回0
Public Sub StopRecvOutput()
    bStopOutput = True
End Sub

'描述:      当类卸载的时候关闭管道，避免目标进程残留后台
Private Sub Class_Terminate()
    Call StopRecvOutput                                                     '停止获取输出
    Call CloseDosIO                                                         '关闭所有管道
End Sub

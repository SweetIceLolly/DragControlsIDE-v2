VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "gdb管道调试测试"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStopRecvOutput 
      Caption         =   "停止获取输出"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdClosePipe 
      Caption         =   "关闭管道"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdBreak 
      Caption         =   "中断程序"
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox edCommand 
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdSendCommand 
      Caption         =   "发送命令"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox edOutput 
      Height          =   5415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   9975
   End
   Begin VB.CommandButton cmdStartPipe 
      Caption         =   "启动管道"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'描述:      示例代码，操作管道类来调用gdb调试TestRes.exe
'作者:      冰棍
'文件:      frmMain.frm
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

'中断指定的进程
Private Declare Function DebugBreakProcess Lib "kernel32" (ByVal hProcess As Long) As Long

'关闭句柄
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

'打开进程
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'创建进程
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'杀掉进程
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

'OpenProcess, dwDesiredAccess
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF                 '进程打开权限

'CreateProcess, dwCreationFlags
Private Const NORMAL_PRIORITY_CLASS = &H20&                         '普通优先级
Private Const CREATE_SUSPENDED = &H4                                '运行之后立即挂起

Dim pi                          As PROCESS_INFORMATION              '进程信息

Private WithEvents GdbPipe      As clsPipe
Attribute GdbPipe.VB_VarHelpID = -1

Private Sub cmdClosePipe_Click()
    GdbPipe.CloseDosIO
End Sub

Private Sub cmdStopRecvOutput_Click()
    GdbPipe.StopRecvOutput
End Sub

Private Sub Form_Load()
    Set GdbPipe = New clsPipe
End Sub

Private Sub cmdSendCommand_Click()
    Dim OutStr  As String
    If GdbPipe.DosInput(Me.edCommand.Text & vbCrLf) = 0 Then
        MsgBox "发送命令失败！"
        Exit Sub
    End If
    GdbPipe.DosOutput OutStr, "(gdb) "
    Me.edOutput.Text = OutStr
    Me.edCommand.Text = ""
    Me.edCommand.SetFocus
End Sub

Private Sub cmdBreak_Click()
    If pi.hProcess <> 0 Then
        If DebugBreakProcess(pi.hProcess) = 0 Then                  '尝试中断目标程序
            MsgBox "挂起指定进程失败！"
        Else
            Dim OutStr  As String
            
            GdbPipe.DosOutput OutStr, "(gdb) "
            Me.edOutput.Text = OutStr
        End If
    Else
        MsgBox "挂起指定进程失败！"
    End If
End Sub

Private Sub cmdStartPipe_Click()
    Dim OutStr  As String
    Dim si      As STARTUPINFO                                      '进程启动信息
    Dim sa      As SECURITY_ATTRIBUTES                              '安全属性
    
    '创建待调试进程
    With sa                                                         '设置安全属性
        .nLength = Len(sa)
        .bInheritHandle = 1                                             '句柄可继承
        .lpSecurityDescriptor = 0
    End With
    If CreateProcess(ByVal 0, "TestRes.exe", sa, sa, ByVal 1, NORMAL_PRIORITY_CLASS Or CREATE_SUSPENDED, ByVal 0, ByVal 0, si, pi) <> 1 Then
        MsgBox "启动进程失败！"
        Exit Sub
    End If
    
    '创建gdb管道
    If GdbPipe.InitDosIO("gdb -q -nw") = 0 Then
        MsgBox "创建管道失败！"
        Exit Sub
    End If
    
    GdbPipe.DosInput "attach " & pi.dwProcessId & vbCrLf            '附加到待调试进程
    GdbPipe.DosInput "continue" & vbCrLf                            '使目标进程继续运行
    
    GdbPipe.DosOutput OutStr, "(gdb) "
    Me.edOutput.Text = OutStr
    
    Me.Caption = "管道已启动！gdb进程ID: " & GdbPipe.dwProcessId & " TestRes.exe进程ID: " & pi.dwProcessId
End Sub

Private Sub edCommand_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSendCommand_Click
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    TerminateProcess pi.hProcess, 0                                 '关掉目标进程
    Set GdbPipe = Nothing
End Sub

'描述:      管道正在输出数据
'参数:      strOutput: 管道中的数据
Private Sub GdbPipe_Output(strOutput As String)
    Me.edOutput.Text = strOutput
    Me.Caption = "DosOutput执行中..."
End Sub

'描述:      管道完成输出数据
Private Sub GdbPipe_OutputFinished()
    Me.Caption = "DosOutput已返回！"
End Sub

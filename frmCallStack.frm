VERSION 5.00
Begin VB.Form frmCallStack 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "调用堆栈"
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DragControlsIDE.DarkListView lvCallStack 
      Height          =   2655
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4683
   End
End
Attribute VB_Name = "frmCallStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'描述:      调用堆栈窗口，在中断状态下显示调用堆栈
'作者:      冰棍
'文件:      frmCallStack.frm
'====================================================

Option Explicit

'定义调用堆栈信息结构
Private Type CallStackInfoStruct
    Address                 As String                                       '地址
    Args                    As String                                       '参数
    File                    As String                                       '文件
    Line                    As Long                                         '行号
End Type

Dim CallStackInfo()         As CallStackInfoStruct                          '所有调用堆栈信息

'描述:      清空所有东西，为下一次调试做准备
Public Sub ClearEverything()
    Me.lvCallStack.Clear
    ReDim CallStackInfo(0)
End Sub

'描述:      获取调用堆栈列表
Public Sub GetCallStack()
    'On Error Resume Next       'todo
    Dim PipeOutput          As String                                       '管道的输出
    Dim OutputLines()       As String                                       '输出的每一行
    Dim StrPos              As Long                                         '查找到的字符串的位置
    Dim BracketLevel        As Long                                         '括号匹配计数，一开始是0，遇到“(”加1, 遇到“)”减1
    Dim NewListItem         As Long                                         '新添加的ListView列表项索引
    Dim i                   As Long
    
    Me.lvCallStack.Clear
    frmMain.DockingPane.Panes(10).Title = Lang_CallStack_Retrieving_Caption
    
    frmMain.GdbPipe.ClearPipe                                                                               '清空管道里的内容
    frmMain.GdbPipe.DosInput "info stack" & vbCrLf                                                          '向gdb发送获取调用堆栈命令
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                                                          '获取gdb输出
    
    OutputLines = Split(PipeOutput, vbCrLf)                                                                 '逐行分割开输出
    ReDim CallStackInfo(UBound(OutputLines) - 1)                                                            '分配信息列表元素
    For i = 0 To UBound(OutputLines)                                                                        '逐行进行分析
        If Trim(OutputLines(i)) <> "(gdb)" Then                                                                 '去掉无用输出“(gdb) ”
            If Mid(OutputLines(i), Len(OutputLines(i))) = vbCr Then                                                 '去掉字符串结尾的换行符
                OutputLines(i) = Left(OutputLines(i), Len(OutputLines(i)) - 1)
            End If
            If OutputLines(i) Like "[#]* * in *(*) at *:\*" Then                                                    '有准确地址；有对应文件
                '例子: #1  0x0040144c in main () at C:\(aa) bb\.cpp:6
                StrPos = InStrRev(OutputLines(i), ":")                                                                  '查找字符串中的“:”，以分割出行号
                CallStackInfo(i).Line = CLng(Right(OutputLines(i), Len(OutputLines(i)) - StrPos))                       '（#n  addr in func (args) at file:[line]）
                OutputLines(i) = Left(OutputLines(i), StrPos - 1)                                                       '（[#n  addr in func (args) at file]:line）
                StrPos = InStrRev(OutputLines(i), ":\")                                                                 '查找字符串中的“:\"，以分割出文件名
                If StrPos = 0 Then                                                                                      '找不到“:\”（新版gdb）就尝试查找“:/”（旧版gdb）
                    StrPos = InStrRev(OutputLines(i), ":/")
                End If
                CallStackInfo(i).File = Right(OutputLines(i), Len(OutputLines(i)) - _
                    InStrRev(OutputLines(i), " at ", StrPos) - 3)                                                       '（#n  addr in func (args) at [file]）
                OutputLines(i) = Left(OutputLines(i), Len(OutputLines(i)) - Len(CallStackInfo(i).File) - 4)             '（[#n  addr in func (args)] at file）
                StrPos = InStr(OutputLines(i), " (")                                                                    '查找字符串中的“ (”，以分割出参数
                CallStackInfo(i).Args = Mid(OutputLines(i), StrPos + 2, Len(OutputLines(i)) - StrPos - 2)               '（#n  addr in func ([args])）
                StrPos = InStr(OutputLines(i), " 0x")                                                                   '查找字符串中的“ 0x”，以分割出地址
                CallStackInfo(i).Address = Mid(OutputLines(i), StrPos + 1, _
                    Len(OutputLines(i)) - StrPos - Len(CallStackInfo(i).Args) - 3)                                      '（#n  [addr in func] (args)）
            ElseIf OutputLines(i) Like "[#]* * in *(*)" Then                                                        '有准确地址；无对应文件
                '例子: #1  0x00403c44 in main ()
                StrPos = InStr(OutputLines(i), " (")                                                                    '查找字符串中的“ (”，以分割出参数
                CallStackInfo(i).Args = Mid(OutputLines(i), StrPos + 2, Len(OutputLines(i)) - StrPos - 2)               '（#n  addr in func ([args])）
                StrPos = InStr(OutputLines(i), " 0x")                                                                   '查找字符串中的“ 0x”，以分割出地址
                CallStackInfo(i).Address = Mid(OutputLines(i), StrPos + 1, _
                    Len(OutputLines(i)) - StrPos - Len(CallStackInfo(i).Args) - 3)                                      '（#n  [addr in func] (args)）
            ElseIf OutputLines(i) Like "[#]* * (*) at *:\*" Then                                                    '无准确地址；有对应文件
                '例子: #0  aaa (a=1, b=2, c=3, d=4) at C:\(aa) bb\.cpp:6
                StrPos = InStrRev(OutputLines(i), ":")                                                                  '查找字符串中的“:”，以分割出行号
                CallStackInfo(i).Line = CLng(Right(OutputLines(i), Len(OutputLines(i)) - StrPos))                       '（#n  func (args) at file:[line]）
                OutputLines(i) = Left(OutputLines(i), StrPos - 1)                                                       '（[#n  func (args) at file]:line）
                StrPos = InStrRev(OutputLines(i), ":\")                                                                 '查找字符串中的“:\"，以分割出文件名
                If StrPos = 0 Then                                                                                      '找不到“:\”（新版gdb）就尝试查找“:/”（旧版gdb）
                    StrPos = InStrRev(OutputLines(i), ":/")
                End If
                CallStackInfo(i).File = Right(OutputLines(i), Len(OutputLines(i)) - _
                    InStrRev(OutputLines(i), " at ", StrPos) - 3)                                                       '（#n  func (args) at [file]）
                OutputLines(i) = Left(OutputLines(i), Len(OutputLines(i)) - Len(CallStackInfo(i).File) - 4)             '（[#n  func (args)] at file）
                StrPos = InStr(OutputLines(i), " ")                                                                     '查找字符串中的“ ”，以去掉开头的序号
                OutputLines(i) = Trim(Right(OutputLines(i), Len(OutputLines(i)) - StrPos))                              '（#n  [func (args)]）
                StrPos = InStr(OutputLines(i), " (")                                                                    '查找字符串中的“ (”，以分割出参数
                CallStackInfo(i).Address = Left(OutputLines(i), StrPos - 1)                                             '（[func] (args)）
                CallStackInfo(i).Args = Mid(OutputLines(i), StrPos + 2, Len(OutputLines(i)) - StrPos - 2)               '（func ([args])）
            ElseIf OutputLines(i) Like "[#]* * (*)" Then                                                            '无准确地址；无对应文件
                '例子: #1  aaa (a=1, b=2, c=3, d=4)
                StrPos = InStr(OutputLines(i), " ")                                                                     '查找字符串中的“ ”，以去掉开头的序号
                OutputLines(i) = Trim(Right(OutputLines(i), Len(OutputLines(i)) - StrPos))                              '（#n  [func (args)]）
                StrPos = InStr(OutputLines(i), " (")                                                                    '查找字符串中的“ (”，以分割出参数
                CallStackInfo(i).Address = Left(OutputLines(i), StrPos - 1)                                             '（[func] (args)）
                CallStackInfo(i).Args = Mid(OutputLines(i), StrPos + 2, Len(OutputLines(i)) - StrPos - 2)               '（func ([args])）
            End If
            
            NewListItem = Me.lvCallStack.AddItem(CallStackInfo(i).Address)                                          '添加新列表项
            If CallStackInfo(i).Args <> "" Then
                Me.lvCallStack.SetItemText CallStackInfo(i).Args, NewListItem, 1
            Else
                CallStackInfo(i).Args = Lang_CallStack_NoArg
                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 1
            End If
            If CallStackInfo(i).File <> "" Then
                Me.lvCallStack.SetItemText CallStackInfo(i).File, NewListItem, 2
            Else
                CallStackInfo(i).File = Lang_CallStack_NoArg
                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 2
            End If
            If CallStackInfo(i).Line <> 0 Then
                Me.lvCallStack.SetItemText CStr(CallStackInfo(i).Line), NewListItem, 3
            Else
                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 3
            End If
        
'            If OutputLines(i) Like "[#]* * in *(*)*" Then                                                           '输出中不带文件名
'
'                OutputLines(i) = Right(OutputLines(i), Len(OutputLines(i)) - Len(Split(OutputLines(i), " ")(0)) - 1)    '（#n func(arg types) (args)）
'                CallStackInfo(i).Address = OutputLines(i)
'                StrPos = InStr(OutputLines(i), "(")                                                                     '查找字符串里第一个“(”
'                CallStackInfo(i).Args = Mid(OutputLines(i), StrPos + 1, Len(OutputLines(i)) - StrPos - 2)               '（#n func(arg types) ([args])）
'                If CallStackInfo(i).Args = "" Then
'                    CallStackInfo(i).Args = Lang_CallStack_NoArg
'                End If
'
'                NewListItem = Me.lvCallStack.AddItem(CallStackInfo(i).Address)                                          '添加新列表项
'                Me.lvCallStack.SetItemText "", NewListItem, 1
'                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 2
'                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 3
'            Else                                                                                                    '输出中带有文件名
'                StrPos = InStrRev(OutputLines(i), ":")                                                                  '（#n func(arg types) (args) at file:line）
'                CallStackInfo(i).Line = CLng(Right(OutputLines(i), Len(OutputLines(i)) - StrPos))                       '（#n func(arg types) (args) at file:[line]）
'                OutputLines(i) = Left(OutputLines(i), StrPos - 1)                                                       '（[#n func(arg types) (args) at file]:line）
'                StrPos = InStrRev(OutputLines(i), ":\")                                                                 '向前查找“:\”（新版gdb）
'                If StrPos = 0 Then                                                                                      '找不到“:\”就查找“:/”（旧版gdb）
'                    StrPos = InStrRev(OutputLines(i), ":/")
'                End If
'                StrPos = InStrRev(OutputLines(i), " at ", StrPos)                                                       '从“:/”的位置继续向前查找“ at ”
'                CallStackInfo(i).File = Replace(Right(OutputLines(i), Len(OutputLines(i)) - StrPos - 3), "/", "\")      '（#n func(arg types) (args) at [file]）
'                OutputLines(i) = Left(OutputLines(i), StrPos - 1)                                                       '（[#n func(arg types) (args)] at file）
'                OutputLines(i) = Right(OutputLines(i), Len(OutputLines(i)) - InStr(OutputLines(i), " ") - 1)            '（#n [func(arg types) (args)]）
'                StrPos = InStr(OutputLines(i), "(")                                                                     '查找字符串里的第一个“(”
'                BracketLevel = 0
'                For StrPos = StrPos + 1 To Len(OutputLines(i))                                                          '往后面查找匹配的“)”（这部分代码与frmLocals的ArrayParser中的代码相似）
'                    If Mid(OutputLines(i), StrPos, 1) = "(" Then                                                            '遇到开括号: 计数+1
'                        BracketLevel = BracketLevel + 1
'                    ElseIf Mid(OutputLines(i), StrPos, 1) = ")" Then                                                        '遇到关括号
'                        If BracketLevel <= 0 Then                                                                               '括号计数为0，即括号已经匹配。此时StrPos是下一个匹配的“)”的位置
'                            CallStackInfo(i).Address = Left(OutputLines(i), StrPos)                                                 '（[func(arg types)] (args)）
'                            Exit For                                                                                                '别继续往后找了
'                        Else                                                                                                    '括号仍未匹配，计数减1，继续往后查找
'                            BracketLevel = BracketLevel - 1
'                        End If
'                    ElseIf Mid(OutputLines(i), StrPos, 1) = """" Then                                                       '遇到“"”，查找到下一个匹配的”"“，确保不会分析到字符串中间去
'                        Do                                                                                                      '一直向后查找“"”，直到不处于字符串中间
'                            StrPos = StrPos + 1
'                        Loop Until (Mid(OutputLines(i), StrPos, 1) = """" And Mid(OutputLines(i), StrPos - 1, 1) <> "\") Or StrPos > Len(OutputLines(i))
'                    End If
'                Next StrPos
'                If StrPos = Len(OutputLines(i)) Then                                                                    '输出里面没有参数
'                    CallStackInfo(i).Args = Lang_CallStack_NoArg                                                            '设置参数为空
'                Else                                                                                                    '输出里面有参数
'                    CallStackInfo(i).Args = Right(OutputLines(i), Len(OutputLines(i)) - StrPos - 1)                         '（func(arg types) [(args)]）
'                End If
'
'                NewListItem = Me.lvCallStack.AddItem(CallStackInfo(i).Address)                                          '添加新列表项
'                Me.lvCallStack.SetItemText CallStackInfo(i).Args, NewListItem, 1
'                Me.lvCallStack.SetItemText CallStackInfo(i).File, NewListItem, 2
'                Me.lvCallStack.SetItemText CStr(CallStackInfo(i).Line), NewListItem, 3
'            End If
        End If
    Next i
    
    frmMain.DockingPane.Panes(10).Title = Lang_CallStack_Caption
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_CallStack_Caption
    
    Me.lvCallStack.Move 0, 0
    
    Me.lvCallStack.AddColumnHeader Lang_Breakpoints_ListViewHeader_Address, 300
    Me.lvCallStack.AddColumnHeader Lang_CallStack_Args, 300
    Me.lvCallStack.AddColumnHeader Lang_Breakpoints_ListViewHeader_File, 150
    Me.lvCallStack.AddColumnHeader Lang_Breakpoints_ListViewHeader_Line
    
    ReDim CallStackInfo(0)                                                  '初始化调用堆栈信息列表
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.lvCallStack.Width = Me.ScaleWidth
    Me.lvCallStack.Height = Me.ScaleHeight
End Sub

Private Sub lvCallStack_Click(iItem As Long, iSubItem As Long, X As Long, Y As Long)
    'On Error Resume Next       'todo
    
    CtlAddToolTip Me.lvCallStack.ListViewHwnd, Lang_Breakpoints_ListViewHeader_Address & ": " & CallStackInfo(iItem).Address & vbCrLf & _
        Lang_CallStack_Args & ": " & CallStackInfo(iItem).Args & vbCrLf & _
        Lang_Breakpoints_ListViewHeader_File & ": " & CallStackInfo(iItem).File & ":" & CallStackInfo(iItem).Line, _
        Lang_CallStack_Tooltip_Title, TTI_INFO
End Sub

Private Sub lvCallStack_DoubleClick(iItem As Long, iSubItem As Long, X As Long, Y As Long)
    On Error Resume Next
    Dim i                   As Long
    
    If CallStackInfo(iItem).File <> "" Then                                                 '如果有对应的文件
        For i = 0 To UBound(CurrentProject.Files)                                               '尝试在工程的文件中查找对应的文件
            If CurrentProject.Files(i).FilePath = CallStackInfo(iItem).File Then                    '查找到对应的文件
                If CurrentProject.Files(i).TargetWindow Is Nothing Then                              '如果有对应的代码窗口就切换过去
                    Dim NewCodeWindow   As frmCodeWindow
                    Dim FileData        As String
                    Dim tmpData         As String
                    
                    Set NewCodeWindow = CreateNewCodeWindow(i)                                              '创建新的代码窗体并设置绑定的文件序号
                    NewCodeWindow.Caption = GetFileName(CallStackInfo(iItem).File)
                    
                    Err.Clear
                    Open CallStackInfo(iItem).File For Input As #1                                          '尝试打开对应的代码文件
                        If Err.Number <> 0 Then
                            Close #1
                            NoSkinMsgBox Lang_Main_Debug_OpenSourceFailure & CallStackInfo(iItem).File, vbExclamation, Lang_Msgbox_Error
                        Else
                            Do While Not EOF(1)
                                Line Input #1, tmpData
                                FileData = FileData & tmpData & vbCrLf
                            Loop
                        End If
                    Close #1
                    
                    frmMain.TabBar.AddForm NewCodeWindow
                Else                                                                                    '没有的话就创建一个新的代码窗口
                    frmMain.TabBar.SwitchToByForm CurrentProject.Files(i).TargetWindow
                End If
                
                CurrentProject.Files(i).TargetWindow.SyntaxEdit.CurrPos.Row = CallStackInfo(iItem).Line
                CurrentProject.Files(i).TargetWindow.SyntaxEdit.SetFocus
                Exit Sub
            End If
        Next i
    End If
End Sub


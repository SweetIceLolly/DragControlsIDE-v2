Attribute VB_Name = "modConfig"
'====================================================
'描述:      提供读写程序配置文件，包括程序设置、语言、用户习惯等函数
'作者:      冰棍
'文件:      modConfig.bas
'====================================================

Option Explicit

'定义代码文件信息结构
Public Type SourceFile
    IsHeaderFile            As Boolean                                          '是否为头文件
    PrevLine                As Long                                             '保存时处在的行号
    Changed                 As Boolean                                          '文件是否被更改
    FilePath                As String                                           '文件路径
    TargetWindow            As frmCodeWindow                                    '对应的代码窗体，每次运行的时候都会不一样
End Type

'定义保存专用的代码文件信息结构
Public Type SourceFile_Save
    IsHeaderFile            As Boolean                                          '是否为头文件
    PrevLine                As Long                                             '保存时处在的行号
    FilePath                As String                                           '文件路径
End Type

'定义工程文件结构
Public Type ProjectFileStruct
    ProjectName             As String                                           '工程名称
    ProjectType             As Integer                                          '工程类型。请见frmMain的ProjectType变量的说明
    Changed                 As Boolean                                          '文件是否被更改
    Files()                 As SourceFile                                       '工程包括的所有文件
End Type

'定义保存专用的工程文件结构
Public Type ProjectFileStruct_Save
    ProjectName             As String                                           '工程名称
    ProjectType             As Integer                                          '工程类型。请见frmMain的ProjectType变量的说明
    Files()                 As SourceFile_Save                                  '工程包括的所有文件
End Type

'定义树视图列表项与文件序号绑定的结构
Public Type TvItemToFileIndex
    TVITEM                  As Long                                             '文件序号对应的树视图列表项
    FileIndex               As Long                                             '树视图列表项对应的文件序号
End Type

'===================================================================
'所有文本变量。由于是多语言，故使用变量来代表每一个出现的字符串
Public Lang_Msgbox_Error                        As String
Public Lang_Msgbox_Confirm                      As String
Public Lang_TitleBar_Max                        As String
Public Lang_TitleBar_Restore                    As String
Public Lang_TitleBar_Min                        As String
Public Lang_TitleBar_Close                      As String
Public Lang_Application_Title                   As String
Public Lang_Breakpoints_Caption                 As String
Public Lang_CallStack_Caption                   As String
Public Lang_CodeWindow_Caption                  As String
Public Lang_ControlBox_Caption                  As String
Public Lang_Create_Caption                      As String
Public Lang_CreateOptions_Caption               As String
Public Lang_Disassembly_Caption                 As String
Public Lang_ErrorList_Caption                   As String
Public Lang_Immediate_Caption                   As String
Public Lang_Locals_Caption                      As String
Public Lang_Memory_Caption                      As String
Public Lang_Modules_Caption                     As String
Public Lang_Output_Caption                      As String
Public Lang_Properties_Caption                  As String
Public Lang_Registers_Caption                   As String
Public Lang_SolutionExplorer_Caption            As String
Public Lang_Threads_Caption                     As String
Public Lang_Watch_Caption                       As String
Public Lang_Create_CreateLabel                  As String
Public Lang_Create_RecentLabel                  As String
Public Lang_Create_NewWindowProgram             As String
Public Lang_Create_NewConsoleProgram            As String
Public Lang_Create_NewEmptyCpp                  As String
Public Lang_Create_OpenProject                  As String
Public Lang_CreateOptions_ProjectNameLabel      As String
Public Lang_CreateOptions_ProjectFolderLabel    As String
Public Lang_CreateOptions_Browse                As String
Public Lang_CreateOptions_Main_NoArgs           As String
Public Lang_CreateOptions_Main_Args             As String
Public Lang_CreateOptions_WinMain               As String
Public Lang_CreateOptions_Include               As String
Public Lang_CreateOptions_OK                    As String
Public Lang_CreateOptions_Cancel                As String
Public Lang_CreateOptions_BrowseCaption         As String
Public Lang_CreateOptions_InvalidProjectPath    As String
Public Lang_CreateOptions_NameConflict_1        As String
Public Lang_CreateOptions_NameConflict_2        As String
Public Lang_CreateOptions_CreationFailure_1     As String
Public Lang_CreateOptions_CreationFailure_2     As String
Public Lang_CreateOptions_SourceFile            As String
Public Lang_CreateOptions_WindowProgram         As String
Public Lang_CreateOptions_ConsoleProgram        As String
Public Lang_CreateOptions_PlainCPP              As String
Public Lang_Main_SaveFailure_1                  As String
Public Lang_Main_saveFailure_2                  As String
Public Lang_Main_SaveBeforeCompile              As String
Public Lang_Main_StartingGcc                    As String
Public Lang_Main_GccStartFailed                 As String
Public Lang_Main_CompileSucceed                 As String
Public Lang_Main_CompileFailed                  As String
Public Lang_Main_RunFailed                      As String
Public Lang_Main_GdbFailed                      As String
Public Lang_Main_GdbAttachFailed_1              As String
Public Lang_Main_GdbAttachFailed_2              As String
Public Lang_Main_GdbLoadSymbolsFailure_1        As String
Public Lang_Main_GdbLoadSymbolsFailure_2        As String
Public Lang_Main_DebugInfo_1                    As String
Public Lang_Main_DebugInfo_2                    As String
Public Lang_SolutionExplorer_RenameFailure_1    As String
Public Lang_SolutionExplorer_RenameFailure_2    As String
'===================================================================

Public CurrentProject       As ProjectFileStruct                                '当前工程的信息
Public ProjectFolderPath    As String                                           '当前工程文件夹的位置（以"\"结尾）
Public ProjectFilePath      As String                                           '当前项目工程文件的位置
Public TvItemBinding()      As TvItemToFileIndex                                '当前工程的TreeView列表项和文件序号的绑定
Public ProjectNameTvItem    As Long                                             'TreeView列表项和工程名称的绑定
Public CodeWindows          As New Collection                                   '当前工程所有的代码窗口
Public IsExiting            As Boolean                                          '当前程序是否正在退出

'描述:      创建一个新的代码窗口，并把它添加到CodeWindows中
'参数:      FileIndex: 代码窗口对应的文件序号
'返回值:    创建的代码窗口
Public Function CreateNewCodeWindow(FileIndex As Long) As frmCodeWindow
    Dim NewCodeWindow       As New frmCodeWindow
    
    NewCodeWindow.FileIndex = FileIndex
    CodeWindows.Add NewCodeWindow, CStr(FileIndex)
    Set CurrentProject.Files(FileIndex).TargetWindow = CodeWindows.Item(CStr(FileIndex))    '文件绑定对应的代码窗口。千万不要绑定到NewCodeWindow！
    Set CreateNewCodeWindow = CodeWindows.Item(CStr(FileIndex))                             '返回创建的代码窗口。千万不要返回NewCodeWindow！
End Function

'描述:      读取对应语言的字符串资源。该函数会通过
'.          提供的第一个资源ID来计算出其他字符串所对应的ID
'参数:      ResID: 对应语言所对应的第一个资源ID。如本程序中文语言所对应的第一个资源ID是1001
'.          LoadMenuTextOnly: 可选，默认为False，如果为True则代表只加载菜单文本。
'.                            因为加载菜单文本会使用到frmMain，frmMain会被加载，
'.                            所以在frmMain的Initialize事件中不宜加载菜单文本，而是应该在Load事件中加载
'返回值:    如果读取成功，返回True；否则返回False
Public Function LoadLanguage(ResID As Long, Optional LoadMenuTextOnly As Boolean = False) As Boolean
    On Error Resume Next
    LoadLanguage = True
    
    '读取菜单字符串
    If LoadMenuTextOnly Then
        Dim id          As Long
        
        For id = 0 To 69
            frmMain.DarkMenu.MenuText(id) = LoadResString(ResID + id)
            If Err.Number <> 0 Then
                LoadLanguage = False
                Exit Function
            End If
        Next id
        Exit Function
    End If
    
    '读取所有的字符串
    Lang_Msgbox_Error = "错误"
    Lang_Msgbox_Confirm = "确认"
    Lang_TitleBar_Max = "最大化"
    Lang_TitleBar_Restore = "还原"
    Lang_TitleBar_Min = "最小化"
    Lang_TitleBar_Close = "关闭"
    Lang_Application_Title = "拖控件大法"
    Lang_Breakpoints_Caption = "断点列表"
    Lang_CallStack_Caption = "调用堆栈"
    Lang_CodeWindow_Caption = "代码窗口"
    Lang_ControlBox_Caption = "控件箱"
    Lang_Create_Caption = "新建项目"
    Lang_CreateOptions_Caption = "新建项目"
    Lang_Disassembly_Caption = "反汇编"
    Lang_ErrorList_Caption = "错误列表"
    Lang_Immediate_Caption = "立即窗口"
    Lang_Locals_Caption = "本地"
    Lang_Memory_Caption = "内存"
    Lang_Modules_Caption = "模块"
    Lang_Output_Caption = "输出"
    Lang_Properties_Caption = "属性"
    Lang_Registers_Caption = "寄存器"
    Lang_SolutionExplorer_Caption = "工程资源管理器"
    Lang_Threads_Caption = "线程"
    Lang_Watch_Caption = "监视窗口"
    Lang_Create_CreateLabel = "创建"
    Lang_Create_RecentLabel = "最近"
    Lang_Create_NewWindowProgram = "       新建窗口程序"
    Lang_Create_NewConsoleProgram = "       新建控制台程序"
    Lang_Create_NewEmptyCpp = "       新建空白C++程序"
    Lang_Create_OpenProject = "       打开工程..."
    Lang_CreateOptions_ProjectNameLabel = "项目名称:"
    Lang_CreateOptions_ProjectFolderLabel = "项目文件夹:"
    Lang_CreateOptions_Browse = "浏览..."
    Lang_CreateOptions_Main_NoArgs = "帮我写好main （无参数）"
    Lang_CreateOptions_Main_Args = "帮我写好main （有参数）"
    Lang_CreateOptions_WinMain = "帮我写好WinMain"
    Lang_CreateOptions_Include = "#include <stdio.h>"
    Lang_CreateOptions_OK = "确定"
    Lang_CreateOptions_Cancel = "取消"
    Lang_CreateOptions_BrowseCaption = "选择项目文件夹"
    Lang_CreateOptions_InvalidProjectPath = "指定的项目文件夹路径无效！"
    Lang_CreateOptions_NameConflict_1 = "检测到同名文件: "
    Lang_CreateOptions_NameConflict_2 = " ，是否继续创建项目？目标文件将会被覆盖！"
    Lang_CreateOptions_CreationFailure_1 = "无法创建"
    Lang_CreateOptions_CreationFailure_2 = " ，请确保项目名称是有效的。"
    Lang_CreateOptions_SourceFile = "源文件"
    Lang_CreateOptions_WindowProgram = "新窗口程序"
    Lang_CreateOptions_ConsoleProgram = "新控制台程序"
    Lang_CreateOptions_PlainCPP = "新空白C++程序"
    Lang_Main_SaveFailure_1 = "无法保存文件："
    Lang_Main_saveFailure_2 = " ，是否继续保存其他文件？"
    Lang_Main_SaveBeforeCompile = "是否先保存所有文件再进行编译？"
    Lang_Main_StartingGcc = "正在启动g++进行编译..."
    Lang_Main_GccStartFailed = "无法启动g++！"
    Lang_Main_CompileSucceed = "编译完成: EXE路径: "
    Lang_Main_CompileFailed = "编译失败！"
    Lang_Main_RunFailed = "无法运行 "
    Lang_Main_GdbFailed = "创建gdb调试管道失败！无法进行调试。"
    Lang_Main_GdbAttachFailed_1 = "gdb附加到进程"
    Lang_Main_GdbAttachFailed_2 = "失败，无法进行调试。"
    Lang_Main_GdbLoadSymbolsFailure_1 = "从可执行文件"
    Lang_Main_GdbLoadSymbolsFailure_2 = " 加载符号失败！这意味着断点、查看本地变量等功能将无法正常工作，是否继续调试？"
    Lang_Main_DebugInfo_1 = "调试正在进行: gdb.exe 进程ID: "
    Lang_Main_DebugInfo_2 = " 进程ID: "
    Lang_SolutionExplorer_RenameFailure_1 = "为文件"
    Lang_SolutionExplorer_RenameFailure_2 = " 重命名失败: "
End Function

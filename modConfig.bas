Attribute VB_Name = "modConfig"
'====================================================
'描述:      提供读写程序配置文件，包括程序设置、语言、用户习惯等函数
'作者:      冰棍
'文件:      modConfig.bas
'====================================================

Option Explicit

'定义cpp文件信息结构
Public Type SourceFile
    IsHeaderFile            As Boolean                                          '是否为头文件
    PrevLine                As Long                                             '保存时处在的行号
    Changed                 As Boolean                                          '文件是否被更改
    FilePath                As String                                           '文件路径
    TargetWindow            As frmCodeWindow                                    '对应的代码窗体，每次运行的时候都会不一样
End Type

'定义工程文件结构
Public Type ProjectFileStruct
    ProjectName             As String                                           '工程名称
    ProjectType             As Integer                                          '工程类型。请见frmMain的ProjectType变量的说明
    Changed                 As Boolean                                          '文件是否被更改
    Files()                 As SourceFile                                       '工程包括的所有文件
End Type

'定义树视图列表项与文件序号绑定的结构
Public Type TvItemToFileIndex
    TVITEM                  As Long                                             '文件序号对应的树视图列表项
    FileIndex               As Long                                             '树视图列表项对应的文件序号
End Type

Public CurrentProject       As ProjectFileStruct                                '当前工程的信息
Public ProjectFolderPath    As String                                           '当前工程文件夹的位置（以"\"结尾）
Public ProjectFilePath      As String                                           '当前项目工程文件的位置
Public TvItemBinding()      As TvItemToFileIndex                                '当前工程的TreeView列表项和文件序号的绑定
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
'返回值:    如果读取成功，返回True；否则返回False
Public Function LoadLanguage(ResID As Long) As Boolean
    On Error Resume Next
    LoadLanguage = True
    
    '读取菜单字符串
    Dim id          As Long
    
    For id = 0 To 69
        frmMain.DarkMenu.MenuText(id) = LoadResString(ResID + id)
        If Err.Number <> 0 Then
            LoadLanguage = False
            Exit Function
        End If
    Next id
End Function

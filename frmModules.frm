VERSION 5.00
Begin VB.Form frmModules 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "模块"
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DragControlsIDE.DarkListView lvModules 
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
   End
End
Attribute VB_Name = "frmModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'描述:      模块窗口，在中断状态下显示调试进程加载的模块
'作者:      冰棍
'文件:      frmModules.frm
'====================================================

Option Explicit

Dim LoadedModuleInfo()  As ModuleInfoStruct                 '已加载模块的信息

'描述:      清空所有东西，为下一次调试做准备
Public Sub ClearEverything()
    Me.lvModules.Clear
    ReDim LoadedModuleInfo(0)
End Sub

'描述:      获取模块列表
Public Sub GetModules()
    'on error resume next
    Dim PipeOutput      As String                                       '管道的输出
    Dim OutputLines()   As String                                       '输出的每一行
    Dim NewListItem     As Long                                         '新添加的ListView列表项索引
    Dim rtnInfo         As ModuleInfoStruct                             '分析得到的模块信息
    Dim i               As Long
    
    Me.lvModules.Clear
    frmMain.DockingPane.Panes(12).Title = Lang_Modules_Caption & Lang_DebugWindow_Retrieving_Caption
    
    frmMain.GdbPipe.ClearPipe                                           '清空管道里的内容
    frmMain.GdbPipe.DosInput "info sharedlibrary" & vbCrLf              '向gdb发送获取模块命令
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) ", 2000                '获取gdb输出
    
    OutputLines = Split(PipeOutput, vbCrLf)                             '逐行分割开输出
    ReDim LoadedModuleInfo(UBound(OutputLines) - 2)                     '分配信息列表元素
    For i = 1 To UBound(OutputLines) - 2                                '逐行进行分析
        If Trim(OutputLines(i)) <> "(gdb)" Then                             '去掉无用输出“(gdb) ”
            rtnInfo = ParseModuleString(OutputLines(i))
            LoadedModuleInfo(i) = rtnInfo
            
            NewListItem = Me.lvModules.AddItem(CStr(i))                         '添加新列表项
            Me.lvModules.SetItemText GetFileName(rtnInfo.File), NewListItem, 1
            Me.lvModules.SetItemText rtnInfo.File, NewListItem, 2
            Me.lvModules.SetItemText rtnInfo.From, NewListItem, 3
            Me.lvModules.SetItemText rtnInfo.To, NewListItem, 4
        End If
    Next i
    
    frmMain.DockingPane(12).Title = Lang_Modules_Caption
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_Modules_Caption
    
    Me.lvModules.Move 0, 0
    
    Me.lvModules.AddColumnHeader "#", 35
    Me.lvModules.AddColumnHeader Lang_Modules_ListViewHeader_FileName, 95
    Me.lvModules.AddColumnHeader Lang_Modules_ListViewHeader_FilePath, 270
    Me.lvModules.AddColumnHeader Lang_Modules_ListViewHeader_From, 90
    Me.lvModules.AddColumnHeader Lang_Modules_ListViewHeader_To, 90
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.lvModules.Width = Me.ScaleWidth
    Me.lvModules.Height = Me.ScaleHeight
End Sub

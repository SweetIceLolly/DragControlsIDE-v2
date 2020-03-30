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
      _extentx        =   6376
      _extenty        =   4683
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
    Dim NewListItem         As Long                                         '新添加的ListView列表项索引
    Dim rtnInfo             As CallStackInfoStruct                          '分析得到的调用堆栈信息
    Dim i                   As Long
    
    Me.lvCallStack.Clear
    frmMain.DockingPane.Panes(10).Title = Lang_CallStack_Retrieving_Caption
    
    frmMain.GdbPipe.ClearPipe                                               '清空管道里的内容
    frmMain.GdbPipe.DosInput "info stack" & vbCrLf                          '向gdb发送获取调用堆栈命令
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                          '获取gdb输出
    
    OutputLines = Split(PipeOutput, vbCrLf)                                 '逐行分割开输出
    ReDim CallStackInfo(UBound(OutputLines) - 1)                            '分配信息列表元素
    For i = 0 To UBound(OutputLines)                                        '逐行进行分析
        If Trim(OutputLines(i)) <> "(gdb)" Then                                 '去掉无用输出“(gdb) ”
            If Mid(OutputLines(i), Len(OutputLines(i))) = vbCr Then                 '去掉字符串结尾的换行符
                OutputLines(i) = Left(OutputLines(i), Len(OutputLines(i)) - 1)
            End If
            
            rtnInfo = ParseCallStackString(OutputLines(i))
            CallStackInfo(i) = rtnInfo
            
            NewListItem = Me.lvCallStack.AddItem(rtnInfo.Address)                   '添加新列表项
            If rtnInfo.Args <> "" Then
                Me.lvCallStack.SetItemText rtnInfo.Args, NewListItem, 1
            Else
                rtnInfo.Args = Lang_CallStack_NoArg
                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 1
            End If
            If rtnInfo.File <> "" Then
                Me.lvCallStack.SetItemText rtnInfo.File, NewListItem, 2
            Else
                rtnInfo.File = Lang_CallStack_NoArg
                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 2
            End If
            If rtnInfo.Line <> 0 Then
                Me.lvCallStack.SetItemText CStr(rtnInfo.Line), NewListItem, 3
            Else
                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 3
            End If
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
    
    If CallStackInfo(iItem).File <> "" Then                                                 '如果有对应的文件
        Dim NewCodeWindow   As frmCodeWindow
        
        '切换到对应的窗口
        Set NewCodeWindow = frmMain.ShowCodeWindow(, CallStackInfo(iItem).File)
        If NewCodeWindow Is Nothing Then
            NoSkinMsgBox Lang_Main_Debug_OpenSourceFailure & CallStackInfo(iItem).File, vbExclamation, Lang_Msgbox_Error
        Else
            NewCodeWindow.SyntaxEdit.CurrPos.Row = CallStackInfo(iItem).Line
            NewCodeWindow.SyntaxEdit.SetFocus
        End If
    End If
End Sub


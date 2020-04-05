VERSION 5.00
Begin VB.Form frmThreads 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "线程"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DragControlsIDE.DarkListView lvThreads 
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
   End
End
Attribute VB_Name = "frmThreads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'描述:      线程窗口，在中断状态下显示调试进程的线程
'作者:      冰棍
'文件:      frmThreads.frm
'====================================================

Option Explicit

Dim ThreadInfo()        As ThreadInfoStruct

'描述:      清空所有东西，为下一次调试做准备
Public Sub ClearEverything()
    Me.lvThreads.Clear
    ReDim ThreadInfo(0)
End Sub

'描述:      获取线程列表
Public Sub GetThreads()
    'on error resume next
    Dim PipeOutput      As String                                       '管道的输出
    Dim OutputLines()   As String                                       '输出的每一行
    Dim NewListItem     As Long                                         '新添加的ListView列表项索引
    Dim rtnInfo         As ThreadInfoStruct                             '分析得到的线程信息
    Dim i               As Long
    
    Me.lvThreads.Clear
    frmMain.DockingPane.Panes(12).Title = Lang_Modules_Caption & Lang_DebugWindow_Retrieving_Caption
    
    frmMain.GdbPipe.ClearPipe                                           '清空管道里的内容
    frmMain.GdbPipe.DosInput "info threads" & vbCrLf                    '向gdb发送获取线程命令
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) ", 2000                '获取gdb输出
    
    OutputLines = Split(PipeOutput, vbCrLf)                             '逐行分割开输出
    ReDim ThreadInfo(UBound(OutputLines) - 1)                           '分配信息列表元素
    For i = 1 To UBound(OutputLines) - 1                                '逐行进行分析
        If Trim(OutputLines(i)) <> "(gdb)" Then                             '去掉无用输出“(gdb) ”
            rtnInfo = ParseThreadString(OutputLines(i))
            ThreadInfo(i) = rtnInfo
            
            NewListItem = Me.lvThreads.AddItem(CStr(i))                         '添加新列表项
            Me.lvThreads.SetItemText GetFileName(rtnInfo.Id), NewListItem, 1
            Me.lvThreads.SetItemText rtnInfo.Frame, NewListItem, 2
            
        End If
    Next i
    
    frmMain.DockingPane(11).Title = Lang_Modules_Caption
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_Threads_Caption
    
    Me.lvThreads.Move 0, 0
    
    Me.lvThreads.AddColumnHeader "#", 35
    Me.lvThreads.AddColumnHeader "ID", 90
    Me.lvThreads.AddColumnHeader Lang_Breakpoints_ListViewHeader_Address, 350
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.lvThreads.Width = Me.ScaleWidth
    Me.lvThreads.Height = Me.ScaleHeight
End Sub

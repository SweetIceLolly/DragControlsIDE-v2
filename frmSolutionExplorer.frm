VERSION 5.00
Begin VB.Form frmSolutionExplorer 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "工程资源管理器"
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DragControlsIDE.DarkMenu mnuItemPopup 
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MENU_ITEM_COUNT =   5
      LEVELS_COUNT    =   5
      LEVELS_2        =   1
      LEVELS_3        =   1
      LEVELS_4        =   1
      LEVELS_5        =   1
      MenuID_1        =   0
      MenuText_1      =   "Popup"
      MenuVisible_1   =   -1  'True
      MenuIcon_1      =   "frmSolutionExplorer.frx":0000
      SUBMENU_ITEM_COUNT_1=   4
      SubMenuID_1_0   =   0
      SubMenuText_1_1 =   "打开(&O)"
      SubMenuID_1_1   =   2
      SubMenuText_1_2 =   "重命名(&R)"
      SubMenuID_1_2   =   3
      SubMenuText_1_3 =   "从项目移除(&E)"
      SubMenuID_1_3   =   4
      SubMenuText_1_4 =   "在文件浏览器中打开(&P)"
      SubMenuID_1_4   =   5
      MenuID_2        =   1
      MenuText_2      =   "打开(&O)"
      MenuVisible_2   =   -1  'True
      MenuIcon_2      =   "frmSolutionExplorer.frx":0018
      SubMenuID_2_0   =   0
      MenuID_3        =   2
      MenuText_3      =   "重命名(&R)"
      MenuVisible_3   =   -1  'True
      MenuIcon_3      =   "frmSolutionExplorer.frx":0030
      SubMenuID_3_0   =   0
      MenuID_4        =   3
      MenuText_4      =   "从项目移除(&E)"
      MenuVisible_4   =   -1  'True
      MenuIcon_4      =   "frmSolutionExplorer.frx":0048
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "在文件浏览器中打开(&P)"
      MenuVisible_5   =   -1  'True
      MenuIcon_5      =   "frmSolutionExplorer.frx":0060
      SubMenuID_5_0   =   0
   End
   Begin DragControlsIDE.DarkMenu mnuProjectItemPopup 
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MENU_ITEM_COUNT =   14
      LEVELS_COUNT    =   14
      LEVELS_2        =   1
      LEVELS_3        =   1
      LEVELS_4        =   2
      LEVELS_5        =   2
      LEVELS_6        =   2
      LEVELS_7        =   2
      LEVELS_8        =   2
      LEVELS_9        =   2
      LEVELS_10       =   2
      LEVELS_11       =   1
      LEVELS_12       =   1
      LEVELS_13       =   1
      LEVELS_14       =   1
      MenuID_1        =   0
      MenuText_1      =   "Popup"
      MenuVisible_1   =   -1  'True
      MenuIcon_1      =   "frmSolutionExplorer.frx":0078
      SUBMENU_ITEM_COUNT_1=   6
      SubMenuID_1_0   =   0
      SubMenuText_1_1 =   "编译工程(&C)"
      SubMenuID_1_1   =   2
      SubMenuText_1_2 =   "添加"
      SubMenuID_1_2   =   3
      SubMenuText_1_3 =   "重命名(&R)"
      SubMenuID_1_3   =   11
      SubMenuText_1_4 =   "删除(&D)"
      SubMenuID_1_4   =   12
      SubMenuText_1_5 =   "在文件浏览器中打开(&O)"
      SubMenuID_1_5   =   13
      SubMenuText_1_6 =   "工程属性(&P)"
      SubMenuID_1_6   =   14
      MenuID_2        =   1
      MenuText_2      =   "编译工程(&C)"
      MenuVisible_2   =   -1  'True
      MenuIcon_2      =   "frmSolutionExplorer.frx":0090
      SubMenuID_2_0   =   0
      MenuID_3        =   2
      MenuText_3      =   "添加"
      MenuVisible_3   =   -1  'True
      MenuIcon_3      =   "frmSolutionExplorer.frx":00A8
      SUBMENU_ITEM_COUNT_3=   7
      SubMenuID_3_0   =   0
      SubMenuText_3_1 =   "文件夹(&F)"
      SubMenuID_3_1   =   4
      SubMenuText_3_2 =   "-"
      SubMenuID_3_2   =   5
      SubMenuText_3_3 =   "窗口(&W)"
      SubMenuID_3_3   =   6
      SubMenuText_3_4 =   "C++文件 (.cpp)"
      SubMenuID_3_4   =   7
      SubMenuText_3_5 =   "C++头文件 (.hpp)"
      SubMenuID_3_5   =   8
      SubMenuText_3_6 =   "C文件 (.c)"
      SubMenuID_3_6   =   9
      SubMenuText_3_7 =   "C头文件 (.h)"
      SubMenuID_3_7   =   10
      MenuID_4        =   3
      MenuText_4      =   "文件夹(&F)"
      MenuVisible_4   =   -1  'True
      MenuIcon_4      =   "frmSolutionExplorer.frx":00C0
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "-"
      MenuVisible_5   =   -1  'True
      MenuIcon_5      =   "frmSolutionExplorer.frx":00D8
      SubMenuID_5_0   =   0
      MenuID_6        =   5
      MenuText_6      =   "窗口(&W)"
      MenuVisible_6   =   -1  'True
      MenuIcon_6      =   "frmSolutionExplorer.frx":00F0
      SubMenuID_6_0   =   0
      MenuID_7        =   6
      MenuText_7      =   "C++文件 (.cpp)"
      MenuVisible_7   =   -1  'True
      MenuIcon_7      =   "frmSolutionExplorer.frx":0108
      SubMenuID_7_0   =   0
      MenuID_8        =   7
      MenuText_8      =   "C++头文件 (.hpp)"
      MenuVisible_8   =   -1  'True
      MenuIcon_8      =   "frmSolutionExplorer.frx":0120
      SubMenuID_8_0   =   0
      MenuID_9        =   8
      MenuText_9      =   "C文件 (.c)"
      MenuVisible_9   =   -1  'True
      MenuIcon_9      =   "frmSolutionExplorer.frx":0138
      SubMenuID_9_0   =   0
      MenuID_10       =   9
      MenuText_10     =   "C头文件 (.h)"
      MenuVisible_10  =   -1  'True
      MenuIcon_10     =   "frmSolutionExplorer.frx":0150
      SubMenuID_10_0  =   0
      MenuID_11       =   10
      MenuText_11     =   "重命名(&R)"
      MenuVisible_11  =   -1  'True
      MenuIcon_11     =   "frmSolutionExplorer.frx":0168
      SubMenuID_11_0  =   0
      MenuID_12       =   11
      MenuText_12     =   "删除(&D)"
      MenuVisible_12  =   -1  'True
      MenuIcon_12     =   "frmSolutionExplorer.frx":0180
      SubMenuID_12_0  =   0
      MenuID_13       =   12
      MenuText_13     =   "在文件浏览器中打开(&O)"
      MenuVisible_13  =   -1  'True
      MenuIcon_13     =   "frmSolutionExplorer.frx":0198
      SubMenuID_13_0  =   0
      MenuID_14       =   13
      MenuText_14     =   "工程属性(&P)"
      MenuVisible_14  =   -1  'True
      MenuIcon_14     =   "frmSolutionExplorer.frx":01B0
      SubMenuID_14_0  =   0
   End
   Begin DragControlsIDE.DarkTreeView SolutionTreeView 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5106
   End
End
Attribute VB_Name = "frmSolutionExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'描述:      工程资源管理器，负责显示工程所包含的目录和文件
'作者:      冰棍
'文件:      frmSolutionExplorer.frm
'====================================================

Option Explicit

'以下三个变量用于创建文件夹
Dim IsCreatingFolder    As Boolean                                  '是否正在创建文件夹
Dim IsCreatingFile      As Boolean                                  '是否正在创建文件
Dim CreatedTreeItem     As Long                                     '正在创建的文件夹或者文件的树视图节点
Dim ParentOfCreating    As Long                                     '正在创建的文件夹或者文件的树视图节点的母节点
Dim CreatingDefaultName As String                                   '正在创建的文件夹或者文件的默认名称

'描述:      递归更新文件夹下的子节点路径
'参数:      ParentIndex: 母文件夹序号
Private Sub RenameFolder(ParentIndex As Long)
    Dim i               As Long
    
    For i = 0 To UBound(CurrentProject.Folders)                     '检查所有文件夹，如果其母文件夹被重命名，那么就更新他的路径
        If CurrentProject.Folders(i).ParentFolder = ParentIndex Then
            CurrentProject.Folders(i).FolderPath = CurrentProject.Folders(ParentIndex).FolderPath & "\" & _
                GetFileName(CurrentProject.Folders(i).FolderPath)
            RenameFolder i                                                  '更新下一层文件夹的路径
        End If
    Next i
    For i = 0 To UBound(CurrentProject.Files)                       '检查所有文件，如果其文件夹被重命名，那么就更新他的路径
        If CurrentProject.Files(i).FolderIndex = ParentIndex Then
            CurrentProject.Files(i).FilePath = ProjectFolderPath & CurrentProject.Folders(ParentIndex).FolderPath & "\" & _
                GetFileName(CurrentProject.Files(i).FilePath)
        End If
    Next i
End Sub

'描述:      “用文件浏览器打开”菜单
Private Sub mnuOpenWithExplorer_Click()
    Dim i               As Long
    Dim hTreeItem       As Long
    
    hTreeItem = Me.SolutionTreeView.GetSelectedItem()
    If hTreeItem = ProjectNameTvItem Then                           '选择的列表项是项目节点
        Shell "explorer.exe /select,""" & ProjectFilePath & """", vbNormalFocus
        Exit Sub
    End If
    For i = 0 To UBound(TvItemBinding)                              '查找列表项对应的文件序号
        If hTreeItem = TvItemBinding(i).TVITEM Then                     '找到对应的文件
            If TvItemBinding(i).IsFolder Then                               '选择的项目是文件夹
                Shell "explorer.exe """ & ProjectFolderPath & CurrentProject.Folders(TvItemBinding(i).FileIndex).FolderPath & """", vbNormalFocus
            Else                                                            '选择的项目是文件
                Shell "explorer.exe /select,""" & CurrentProject.Files(TvItemBinding(i).FileIndex).FilePath & """", vbNormalFocus
            End If
        End If
    Next i
End Sub

'描述:      “新建文件夹”菜单
Private Sub mnuCreateFolder_Click()
    Dim hTreeItem       As Long
    
    IsCreatingFolder = True                                                         '标记为正在创建文件夹
    IsCreatingFile = False
    CreatingDefaultName = Lang_SolutionExplorer_NewFolderName
    hTreeItem = Me.SolutionTreeView.GetSelectedItem()
    CreatedTreeItem = Me.SolutionTreeView.AddItem(CreatingDefaultName, hTreeItem)   '创建文件夹节点
    Me.SolutionTreeView.ExpandItems hTreeItem, 2
    ParentOfCreating = hTreeItem
    Me.SolutionTreeView.EditLabel CreatedTreeItem                                   '开始编辑标签
End Sub

'描述:      添加文件过程
'参数:      FileName: 添加的文件名
Private Sub mnuAddFile_Click(FileName As String)
    Dim hTreeItem       As Long
    
    IsCreatingFile = True                                                           '标记为正在创建文件
    IsCreatingFolder = False
    CreatingDefaultName = FileName                                                  '设置默认名称
    hTreeItem = Me.SolutionTreeView.GetSelectedItem()
    CreatedTreeItem = Me.SolutionTreeView.AddItem(FileName, hTreeItem)              '创建文件节点
    Me.SolutionTreeView.ExpandItems hTreeItem, 2
    ParentOfCreating = hTreeItem
    Me.SolutionTreeView.EditLabel CreatedTreeItem                                   '开始编辑标签
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_SolutionExplorer_Caption
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.SolutionTreeView.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub mnuItemPopup_MenuItemClicked(MenuID As Integer)
    Me.mnuItemPopup.HideMenu
    Select Case MenuID
        Case 1                                  '打开
            Call SolutionTreeView_DoubleClick(1, 0, 0, 0)
        
        Case 2                                  '重命名
            Me.SolutionTreeView.EditLabel Me.SolutionTreeView.GetSelectedItem()
        
        Case 3                                  '从项目移除
        
        Case 4                                  '用文件浏览器打开
            Call mnuOpenWithExplorer_Click
        
    End Select
End Sub

Private Sub mnuProjectItemPopup_MenuItemClicked(MenuID As Integer)
    Me.mnuProjectItemPopup.HideMenu
    Select Case MenuID
        Case 1                                  '编译工程
        
        Case 3                                  '文件夹
            Call mnuCreateFolder_Click
        
        Case 5                                  '添加窗口
            
        Case 6                                  '添加cpp
            Call mnuAddFile_Click("新cpp文件.cpp")
        
        Case 7                                  '添加hpp
            Call mnuAddFile_Click("新hpp文件.hpp")
        
        Case 8                                  '添加c
            Call mnuAddFile_Click("新c文件.c")
        
        Case 9                                  '添加h
            Call mnuAddFile_Click("新h文件.h")
        
        Case 10                                 '重命名
            Me.SolutionTreeView.EditLabel Me.SolutionTreeView.GetSelectedItem()
        
        Case 11                                 '删除
            
        
        Case 12                                 '用文件浏览器打开
            Call mnuOpenWithExplorer_Click
        
        Case 13                                 '工程属性
        
    End Select
End Sub

Public Sub SolutionTreeView_BeginLabelEdit(ByVal hTreeItem As Long, bCancel As Boolean)
    Dim i               As Long
    
    If IsCreatingFolder Then                                                                    '如果正在创建文件夹，则允许更改
        Exit Sub
    End If
    If IsCreatingFile Then                                                                      '如果正在创建文件，则自动选取小数点前面的文本
        GoTo SelectEditboxText
    End If
    
    If hTreeItem = ProjectNameTvItem Then                                                       '如果列表项对应的是工程名称，则允许更改
        Exit Sub
    End If
    For i = 0 To UBound(TvItemBinding)                                                          '查找列表项对应的文件序号
        If hTreeItem = TvItemBinding(i).TVITEM Then                                                 '如果能找到对应的文件，说明选择的列表项是文件而不是项目节点
            If TvItemBinding(i).IsFolder Then                                                           '选择的项目是文件夹，允许修改
                Exit Sub
            Else                                                                                        '选择的项目是文件
                GoTo SelectEditboxText
            End If
            Exit Sub
        End If
    Next i
    bCancel = True                                                                              '如果找不到对应的文件，说明选择的列表项是不允许重命名的项目节点
    
SelectEditboxText:
    Dim hwndLabelEditBox    As Long                                                             '进行标签编辑的文本框句柄
    Dim LabelStr            As String                                                           '当前准备编辑的标签的文本
    Dim DotPos              As Integer                                                          '小数点“.”在标签文本里的位置
    
    '如果标签里面有“.”，那么只选择“.”前面的文本
    LabelStr = Me.SolutionTreeView.GetItemText(hTreeItem)                                       '获取当前准备编辑的标签的文本
    DotPos = InStrRev(LabelStr, ".")                                                            '在文本中查找“.”
    If DotPos <> 0 Then                                                                         '如果找到小数点
        hwndLabelEditBox = SendMessageA(Me.SolutionTreeView.TreeViewHwnd, TVM_GETEDITCONTROL, 0, 0) '获取进行标签编辑的文本框句柄
        SetPropA hwndLabelEditBox, "PrevWndProc", _
            SetWindowLongA(hwndLabelEditBox, GWL_WNDPROC, AddressOf TreeViewEditBoxWindowProc)      '设置标签编辑的文本框的子类化，处理选择文本的消息
        SetPropA hwndLabelEditBox, "DotPos", ByVal DotPos - 1                                       '记录“.”的位置，以便文本框的子类化修改选择的文本
    End If
End Sub

Public Sub SolutionTreeView_Click(bCancel As Boolean)
    
End Sub

Public Sub SolutionTreeView_DoubleClick(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    On Error Resume Next
    Dim CurrSelItem     As Long
    Dim i               As Long
    
    CurrSelItem = Me.SolutionTreeView.GetSelectedItem()                                         '获取选择的树视图列表项
    For i = 0 To UBound(TvItemBinding)                                                          '查找列表项对应的文件序号
        If CurrSelItem = TvItemBinding(i).TVITEM Then
            If TvItemBinding(i).IsFolder Then                                                           '如果选择的项目是文件夹
                Me.SolutionTreeView.ExpandItems CurrSelItem, 3                                              '切换节点展开状态
                Me.SolutionTreeView.EndEditLabel False                                                      '取消编辑标签
            Else                                                                                        '如果选择的项目是代码文件
                If CurrentProject.Files(TvItemBinding(i).FileIndex).TargetWindow Is Nothing Then
                    Dim NewCodeWindow   As frmCodeWindow                                                        '新建的代码框窗体
                    
                    Set NewCodeWindow = CreateNewCodeWindow(TvItemBinding(i).FileIndex)                         '创建新的代码窗体并设置绑定的文件序号
                    NewCodeWindow.Caption = GetFileName(CurrentProject.Files(TvItemBinding(i).FileIndex).FilePath)
                    frmMain.TabBar.AddForm NewCodeWindow
                Else
                    frmMain.TabBar.SwitchToByForm CurrentProject.Files(TvItemBinding(i).FileIndex).TargetWindow '切换到对应的窗口
                    CurrentProject.Files(TvItemBinding(i).FileIndex).TargetWindow.SyntaxEdit.SetFocus
                End If
            End If
            Exit For
        End If
    Next i
End Sub

Public Sub SolutionTreeView_EndLabelEdit(ByVal hTreeItem As Long, NewText As String, bCancel As Boolean)
    If NewText = vbNullChar Or NewText = "" Then                                                '如果NewText为vbNullChar，则说明编辑被取消了
        If IsCreatingFolder Or IsCreatingFile Then                                                  '如果正在创建文件夹或者文件就使用默认名称
            NewText = CreatingDefaultName
        Else                                                                                        '不是创建文件夹的话就取消重命名
            Exit Sub
        End If
    End If
    
    If NewText = "" Then                                                                        '是空文本
        If IsCreatingFolder Or IsCreatingFile Then                                                  '如果正在创建文件或者文件夹就取消创建
            IsCreatingFolder = False
            IsCreatingFile = False
            Me.SolutionTreeView.RemoveItem CreatedTreeItem
        End If
        Exit Sub
    End If
    
    '尝试进行重命名
    On Error Resume Next
    Dim i   As Long
    
    If Not CheckInvalidFileName(NewText) Then                                                   '检查非法文件名
        If IsCreatingFolder Or IsCreatingFile Then                                                  '如果正在创建文件或者文件夹就取消创建
            IsCreatingFolder = False
            IsCreatingFile = False
            Me.SolutionTreeView.RemoveItem CreatedTreeItem
        End If
        NoSkinMsgBox Lang_SolutionExplorer_InvalidName, vbExclamation, Lang_Msgbox_Error
        bCancel = True
        Exit Sub
    End If
    
    If IsCreatingFolder Or IsCreatingFile Then                                                      '如果正在创建文件或者文件夹
        Dim FolderPath          As String                                                           '创建文件或者文件夹的位置
        Dim ParentFolderIndex   As Long                                                             '创建的文件或者文件夹的节点的母节点的索引
        
        For i = 0 To UBound(TvItemBinding)                                                          '查找对应的母节点在CurrentProject.Folders的索引
            If ParentOfCreating = TvItemBinding(i).TVITEM Then                                          '记录下母节点所匹配的索引并获取母节点路径，形成完整的相对路径
                ParentFolderIndex = TvItemBinding(i).FileIndex
                FolderPath = CurrentProject.Folders(ParentFolderIndex).FolderPath & "\" & FolderPath
                Exit For
            End If
        Next i
        
        Err.Clear
        If IsCreatingFolder Then                                                                    '如果是创建文件夹
            MkDir ProjectFolderPath & FolderPath & NewText
        Else
            If Dir(ProjectFolderPath & FolderPath & NewText, vbDirectory Or vbNormal Or vbReadOnly Or vbHidden Or vbSystem) = "" Then
                Open ProjectFolderPath & FolderPath & NewText For Binary As #1
                Close #1
            Else                                                                                    '检测到重名
                Err.Raise 75
            End If
        End If
        If Err.Number <> 0 Then                                                                     '创建文件夹时发生错误
            Me.SolutionTreeView.RemoveItem CreatedTreeItem
            MsgBox "创建失败！", vbExclamation, Lang_Msgbox_Error       'todo
        Else                                                                                        '创建文件或者文件夹成功
            CurrentProject.Changed = True                                                               '标记工程文件为已更改
            If IsCreatingFolder Then                                                                    '如果是创建文件夹就更新项目信息里的文件夹信息
                ReDim Preserve CurrentProject.Folders(UBound(CurrentProject.Folders) + 1)                   '添加项目信息里的文件夹信息
                ReDim Preserve TvItemBinding(UBound(TvItemBinding) + 1)                                     '添加树视图项目绑定
                With TvItemBinding(UBound(TvItemBinding))                                                   '设置树视图项目绑定
                    .FileIndex = UBound(CurrentProject.Folders)                                                 '文件夹索引
                    .TVITEM = CreatedTreeItem                                                                   '树视图节点
                    .IsFolder = True                                                                            '标记为文件夹
                End With
                With CurrentProject.Folders(UBound(CurrentProject.Folders))                                 '设置项目信息里的文件夹信息
                    .FolderPath = FolderPath & NewText                                                          '文件夹路径
                    If ParentOfCreating = ProjectNameTvItem Then                                                '如果母节点是项目节点 就把索引设置为0（即在项目目录下）
                        .ParentFolder = 0
                    Else                                                                                        '否则就记录母节点
                        .ParentFolder = TvItemBinding(i).FileIndex
                    End If
                End With
            Else                                                                                        '如果是创建文件就更新项目信息里的文件信息
                ReDim Preserve CurrentProject.Files(UBound(CurrentProject.Files) + 1)                       '添加项目信息里的文件信息
                ReDim Preserve TvItemBinding(UBound(TvItemBinding) + 1)                                     '添加树视图项目绑定
                With TvItemBinding(UBound(TvItemBinding))                                                   '设置树视图项目绑定
                    .FileIndex = UBound(CurrentProject.Files)                                                   '文件索引
                    .TVITEM = CreatedTreeItem                                                                   '树视图节点
                    .IsFolder = False                                                                           '标记为文件
                End With
                With CurrentProject.Files(UBound(CurrentProject.Files))                                     '设置项目信息里的文件信息
                    .FilePath = ProjectFolderPath & FolderPath & NewText                                        '文件路径
                    If ParentOfCreating = ProjectNameTvItem Then                                                '如果母节点是项目节点 就把索引设置为0（即在项目目录下）
                        .FolderIndex = 0
                    Else                                                                                        '否则就记录母节点
                        .FolderIndex = ParentFolderIndex
                    End If
                    .Changed = False                                                                            '标记文件为未更改
                    .PrevLine = 0
                    ReDim .Breakpoints(0)                                                                       '初始化文件断点列表
                End With
            End If
            
            IsCreatingFolder = False
            IsCreatingFile = False
        End If
        Exit Sub
    End If
    
    If hTreeItem = ProjectNameTvItem Then                                                       '如果列表项对应的是工程名称，则更改工程文件名
        Err.Clear
        Name ProjectFilePath As ProjectFolderPath & NewText & ".myproj"
        If Err.Number <> 0 Then                                                                     '重命名时发生错误
            NoSkinMsgBox Lang_SolutionExplorer_RenameFailure_1 & ProjectFilePath & _
                Lang_SolutionExplorer_RenameFailure_2 & Err.Number & " - " & Err.Description, vbExclamation, Lang_Msgbox_Error
            bCancel = True
        Else                                                                                        '重命名成功
            ProjectFilePath = ProjectFolderPath & NewText & ".myproj"                                   '更新工程文件路径
            CurrentProject.ProjectName = NewText                                                        '更新工程名称
            frmMain.Caption = NewText & " - " & Lang_Application_Title                                  '更改主窗口标题
            CurrentProject.Changed = True                                                               '标记工程已更改
        End If
        Exit Sub
    End If
    For i = 0 To UBound(TvItemBinding)                                                          '查找列表项对应的文件序号
        If hTreeItem = TvItemBinding(i).TVITEM Then                                                 '找到匹配的文件就进行重命名
            If TvItemBinding(i).IsFolder Then                                                           '如果选择的项目是文件夹
                With CurrentProject.Folders(TvItemBinding(i).FileIndex)
                    Err.Clear
                    If .ParentFolder = 0 Then                                                                    '如果是在根目录下，就不需要在路径中加“\”
                        Name ProjectFolderPath & .FolderPath As ProjectFolderPath & NewText
                    Else
                        Name ProjectFolderPath & .FolderPath As ProjectFolderPath & CurrentProject.Folders(.ParentFolder).FolderPath & "\" & NewText
                    End If
                    If Err.Number <> 0 Then                                                                     '重命名时发生错误
                        MsgBox "Error!"     'todo
                        bCancel = True
                    Else                                                                                        '重命名成功
                        If .ParentFolder = 0 Then                                                               '更新相对路径
                            .FolderPath = NewText
                        Else
                            .FolderPath = CurrentProject.Folders(.ParentFolder).FolderPath & "\" & NewText
                        End If
                        RenameFolder TvItemBinding(i).FileIndex                                                     '更新这个节点下所有子节点的路径
                    End If
                End With
            Else                                                                                        '如果选择的项目是文件
                With CurrentProject.Files(TvItemBinding(i).FileIndex)
                    Err.Clear
                    If .FolderIndex = 0 Then                                                                    '如果是在根目录下，就不需要在路径中加“\”
                        Name (.FilePath) As ProjectFolderPath & NewText
                    Else
                        Name (.FilePath) As ProjectFolderPath & CurrentProject.Folders(.FolderIndex).FolderPath & "\" & NewText
                    End If
                    If Err.Number <> 0 Then                                                                     '重命名时发生错误
                        NoSkinMsgBox Lang_SolutionExplorer_RenameFailure_1 & .FilePath & _
                            Lang_SolutionExplorer_RenameFailure_2 & Err.Number & " - " & Err.Description, vbExclamation, Lang_Msgbox_Error
                        bCancel = True
                    Else                                                                                        '重命名成功
                        .TargetWindow.Caption = NewText                                                         '刷新窗口标题
                        frmMain.TabBar.UpdateCaptions
                        If .FolderIndex = 0 Then                                                                '更新文件路径
                            Name (.FilePath) As ProjectFolderPath & NewText
                        Else
                            Name (.FilePath) As ProjectFolderPath & CurrentProject.Folders(.FolderIndex).FolderPath & "\" & NewText
                        End If
                        .FilePath = ProjectFolderPath & CurrentProject.Folders(.FolderIndex).FolderPath & "\" & NewText
                    End If
                End With
            End If
            Exit Sub
        End If
    Next i
    bCancel = True                                                                              '其实应该不会找不到对应的文件，但是如果真的找不到就取消操作吧
End Sub

Public Sub SolutionTreeView_ItemExpanding(ByVal hTreeItem As Long, bCancel As Boolean)

End Sub

Public Sub SolutionTreeView_KeyDown(ByVal KeyCode As Long, ByVal IsLongPress As Boolean)
    If KeyCode = vbKeyF2 Then                                                                   '响应F2键: 重命名
        Me.SolutionTreeView.EditLabel Me.SolutionTreeView.GetSelectedItem()
    ElseIf KeyCode = VK_APPS Then                                                               '响应菜单键: 弹出菜单
        Call SolutionTreeView_RightClick(True)
    End If
End Sub

Public Sub SolutionTreeView_KeyUp(ByVal KeyCode As Long)

End Sub

Public Sub SolutionTreeView_MouseDown(ByVal Button As Long, ByVal X As Long, ByVal Y As Long)
    Me.SolutionTreeView.SelectItem Me.SolutionTreeView.HitTest(X, Y)                            '选择鼠标按下的位置的列表项
End Sub

Public Sub SolutionTreeView_MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Public Sub SolutionTreeView_MouseUp(ByVal Button As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

'参数:      bPopupMenuFromItem: 是否根据列表项的位置弹出菜单。用于处理菜单键
Public Sub SolutionTreeView_RightClick(bPopupMenuFromItem As Boolean)
    Dim i               As Long
    Dim hTreeItem       As Long
    Dim ItemRect        As RECT
    Dim WindowRect      As RECT
    
    hTreeItem = Me.SolutionTreeView.GetSelectedItem()
    If bPopupMenuFromItem Then                                                                  '如果是根据根据列表项的位置弹出菜单，就获取列表项的位置
        CopyMemory ItemRect, hTreeItem, ByVal 4                                                     '*(HTREEITEM*)&ItemRect = hTreeItem
        SendMessageA Me.SolutionTreeView.TreeViewHwnd, TVM_GETITEMRECT, ByVal 0, ByVal VarPtr(ItemRect)
        GetWindowRect Me.SolutionTreeView.TreeViewHwnd, WindowRect
        ItemRect.Left = WindowRect.Left * Screen.TwipsPerPixelX                                     '计算出列表项相对于屏幕上的坐标
        ItemRect.bottom = (ItemRect.bottom + WindowRect.Top) * Screen.TwipsPerPixelY
    Else                                                                                        '否则使用菜单的默认弹出位置
        ItemRect.Left = -1
        ItemRect.bottom = -1
    End If
    
    '判断选定的列表项的类型并弹出对应的菜单
    If hTreeItem = 0 Then
        Exit Sub
    End If
    If CurrentProject.ProjectType = 1 Then                                                      '如果工程类型不是窗口程序就不允许添加窗口
        Me.mnuProjectItemPopup.MenuEnabled(5) = True
    Else
        Me.mnuProjectItemPopup.MenuEnabled(5) = False
    End If
    frmMain.DarkTitleBar.SetFocus
    If hTreeItem = ProjectNameTvItem Then                                                       '如果选择的项目是工程文件
        Me.mnuProjectItemPopup.PopupMenu 0, CSng(ItemRect.Left), CSng(ItemRect.bottom)
    Else                                                                                        '否则检查选择的项目是否为代码文件
        For i = 0 To UBound(TvItemBinding)
            If hTreeItem = TvItemBinding(i).TVITEM Then                                             '如果能找到对应的文件，说明选择的列表项是文件而不是项目节点
                If TvItemBinding(i).IsFolder Then                                                       '如果选择的列表项是文件夹
                    Me.mnuProjectItemPopup.PopupMenu 0, CSng(ItemRect.Left), CSng(ItemRect.bottom)
                    Exit Sub
                Else                                                                                    '如果选择的列表项是文件
                    Me.mnuItemPopup.PopupMenu 0, CSng(ItemRect.Left), CSng(ItemRect.bottom)
                    Exit Sub
                End If
            End If
        Next i
    End If
End Sub

Public Sub SolutionTreeView_SelChanged(ByVal hOldTreeItem As Long, ByVal hNewTreeItem As Long)
    
End Sub

Public Sub SolutionTreeView_SelChanging(ByVal hOldTreeItem As Long, ByVal hNewTreeItem As Long, bCancel As Boolean)
    
End Sub

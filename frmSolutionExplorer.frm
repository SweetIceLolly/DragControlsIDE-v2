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
      MENU_ITEM_COUNT =   7
      LEVELS_COUNT    =   7
      LEVELS_2        =   1
      LEVELS_3        =   1
      LEVELS_4        =   1
      LEVELS_5        =   1
      LEVELS_6        =   1
      LEVELS_7        =   1
      MenuID_1        =   0
      MenuText_1      =   "Popup"
      MenuVisible_1   =   -1  'True
      MenuIcon_1      =   "frmSolutionExplorer.frx":0000
      SUBMENU_ITEM_COUNT_1=   6
      SubMenuID_1_0   =   0
      SubMenuText_1_1 =   "打开(&O)"
      SubMenuID_1_1   =   2
      SubMenuText_1_2 =   "重命名(&R)"
      SubMenuID_1_2   =   3
      SubMenuText_1_3 =   "复制(&C)"
      SubMenuID_1_3   =   4
      SubMenuText_1_4 =   "从项目移除(&E)"
      SubMenuID_1_4   =   5
      SubMenuText_1_5 =   "删除(&D)"
      SubMenuID_1_5   =   6
      SubMenuText_1_6 =   "在文件浏览器中打开(&P)"
      SubMenuID_1_6   =   7
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
      MenuText_4      =   "复制(&C)"
      MenuVisible_4   =   -1  'True
      MenuIcon_4      =   "frmSolutionExplorer.frx":0048
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "从项目移除(&E)"
      MenuVisible_5   =   -1  'True
      MenuIcon_5      =   "frmSolutionExplorer.frx":0060
      SubMenuID_5_0   =   0
      MenuID_6        =   5
      MenuText_6      =   "删除(&D)"
      MenuVisible_6   =   -1  'True
      MenuIcon_6      =   "frmSolutionExplorer.frx":0078
      SubMenuID_6_0   =   0
      MenuID_7        =   6
      MenuText_7      =   "在文件浏览器中打开(&P)"
      MenuVisible_7   =   -1  'True
      MenuIcon_7      =   "frmSolutionExplorer.frx":0090
      SubMenuID_7_0   =   0
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
      MENU_ITEM_COUNT =   13
      LEVELS_COUNT    =   13
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
      MenuID_1        =   0
      MenuText_1      =   "Popup"
      MenuVisible_1   =   -1  'True
      MenuIcon_1      =   "frmSolutionExplorer.frx":00A8
      SUBMENU_ITEM_COUNT_1=   5
      SubMenuID_1_0   =   0
      SubMenuText_1_1 =   "编译工程(&C)"
      SubMenuID_1_1   =   2
      SubMenuText_1_2 =   "添加"
      SubMenuID_1_2   =   3
      SubMenuText_1_3 =   "重命名(&R)"
      SubMenuID_1_3   =   11
      SubMenuText_1_4 =   "在文件浏览器中打开(&O)"
      SubMenuID_1_4   =   12
      SubMenuText_1_5 =   "工程属性(&P)"
      SubMenuID_1_5   =   13
      MenuID_2        =   1
      MenuText_2      =   "编译工程(&C)"
      MenuVisible_2   =   -1  'True
      MenuIcon_2      =   "frmSolutionExplorer.frx":00C0
      SubMenuID_2_0   =   0
      MenuID_3        =   2
      MenuText_3      =   "添加"
      MenuVisible_3   =   -1  'True
      MenuIcon_3      =   "frmSolutionExplorer.frx":00D8
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
      MenuIcon_4      =   "frmSolutionExplorer.frx":00F0
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "-"
      MenuVisible_5   =   -1  'True
      MenuIcon_5      =   "frmSolutionExplorer.frx":0108
      SubMenuID_5_0   =   0
      MenuID_6        =   5
      MenuText_6      =   "窗口(&W)"
      MenuVisible_6   =   -1  'True
      MenuIcon_6      =   "frmSolutionExplorer.frx":0120
      SubMenuID_6_0   =   0
      MenuID_7        =   6
      MenuText_7      =   "C++文件 (.cpp)"
      MenuVisible_7   =   -1  'True
      MenuIcon_7      =   "frmSolutionExplorer.frx":0138
      SubMenuID_7_0   =   0
      MenuID_8        =   7
      MenuText_8      =   "C++头文件 (.hpp)"
      MenuVisible_8   =   -1  'True
      MenuIcon_8      =   "frmSolutionExplorer.frx":0150
      SubMenuID_8_0   =   0
      MenuID_9        =   8
      MenuText_9      =   "C文件 (.c)"
      MenuVisible_9   =   -1  'True
      MenuIcon_9      =   "frmSolutionExplorer.frx":0168
      SubMenuID_9_0   =   0
      MenuID_10       =   9
      MenuText_10     =   "C头文件 (.h)"
      MenuVisible_10  =   -1  'True
      MenuIcon_10     =   "frmSolutionExplorer.frx":0180
      SubMenuID_10_0  =   0
      MenuID_11       =   10
      MenuText_11     =   "重命名(&R)"
      MenuVisible_11  =   -1  'True
      MenuIcon_11     =   "frmSolutionExplorer.frx":0198
      SubMenuID_11_0  =   0
      MenuID_12       =   11
      MenuText_12     =   "在文件浏览器中打开(&O)"
      MenuVisible_12  =   -1  'True
      MenuIcon_12     =   "frmSolutionExplorer.frx":01B0
      SubMenuID_12_0  =   0
      MenuID_13       =   12
      MenuText_13     =   "工程属性(&P)"
      MenuVisible_13  =   -1  'True
      MenuIcon_13     =   "frmSolutionExplorer.frx":01C8
      SubMenuID_13_0  =   0
   End
   Begin DragControlsIDE.DarkTreeView SolutionTreeView 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
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

Private Sub Form_Load()
    Me.Caption = Lang_SolutionExplorer_Caption
    
    'ToDo: Set context menu text
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
        
        Case 3                                  '复制
        
        Case 4                                  '从项目移除
        
        Case 5                                  '删除
        
        Case 6                                  '用文件浏览器打开
            Dim i           As Long
            Dim hTreeItem    As Long
            
            hTreeItem = Me.SolutionTreeView.GetSelectedItem()
            For i = 0 To UBound(TvItemBinding)                              '查找列表项对应的文件序号
                If hTreeItem = TvItemBinding(i).TVITEM Then                     '找到对应的文件
                    Shell "explorer.exe /select,""" & CurrentProject.Files(TvItemBinding(i).FileIndex).FilePath & """", vbNormalFocus
                End If
            Next i
        
    End Select
End Sub

Private Sub mnuProjectItemPopup_MenuItemClicked(MenuID As Integer)
    Me.mnuProjectItemPopup.HideMenu
    Select Case MenuID
        Case 1                                  '编译工程
            
        
        Case 3                                  '文件夹
        
        Case 5                                  '添加窗口
            
        Case 6                                  '添加cpp
        
        Case 7                                  '添加hpp
        
        Case 8                                  '添加c
        
        Case 9                                  '添加h
        
        Case 10                                 '重命名
            Me.SolutionTreeView.EditLabel Me.SolutionTreeView.GetSelectedItem()
        
        Case 11                                 '用文件浏览器打开
            Shell "explorer.exe /select,""" & ProjectFilePath & """", vbNormalFocus
        
        Case 12                                 '工程属性
        
    End Select
End Sub

Public Sub SolutionTreeView_BeginLabelEdit(ByVal hTreeItem As Long, bCancel As Boolean)
    Dim i               As Long
    
    If hTreeItem = ProjectNameTvItem Then                                                       '如果列表项对应的是工程名称，则也允许更改
        Exit Sub
    End If
    For i = 0 To UBound(TvItemBinding)                                                          '查找列表项对应的文件序号
        If hTreeItem = TvItemBinding(i).TVITEM Then                                                 '如果能找到对应的文件，说明选择的列表项是文件而不是项目节点
            '如果标签里面有“.”，那么只选择“.”前面的文本
            Dim hwndLabelEditBox    As Long                                                             '进行标签编辑的文本框句柄
            Dim LabelStr            As String                                                           '当前准备编辑的标签的文本
            Dim DotPos              As Integer                                                          '小数点“.”在标签文本里的位置
            
            LabelStr = Me.SolutionTreeView.GetItemText(hTreeItem)                                       '获取当前准备编辑的标签的文本
            DotPos = InStrRev(LabelStr, ".")                                                            '在文本中查找“.”
            If DotPos <> 0 Then                                                                         '如果找到小数点
                hwndLabelEditBox = SendMessageA(Me.SolutionTreeView.TreeViewHwnd, TVM_GETEDITCONTROL, 0, 0) '获取进行标签编辑的文本框句柄
                SetPropA hwndLabelEditBox, "PrevWndProc", _
                    SetWindowLongA(hwndLabelEditBox, GWL_WNDPROC, AddressOf TreeViewEditBoxWindowProc)      '设置标签编辑的文本框的子类化，处理选择文本的消息
                SetPropA hwndLabelEditBox, "DotPos", ByVal DotPos - 1                                       '记录“.”的位置，以便文本框的子类化修改选择的文本
            End If
            
            Exit Sub
        End If
    Next i
    bCancel = True                                                                              '如果找不到对应的文件，说明选择的列表项是不允许重命名的项目节点
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
            If CurrentProject.Files(TvItemBinding(i).FileIndex).TargetWindow Is Nothing Then
                Dim NewCodeWindow   As frmCodeWindow                                                        '新建的代码框窗体
                
                Set NewCodeWindow = CreateNewCodeWindow(TvItemBinding(i).FileIndex)                         '创建新的代码窗体并设置绑定的文件序号
                NewCodeWindow.Caption = GetFileName(CurrentProject.Files(TvItemBinding(i).FileIndex).FilePath)
                frmMain.TabBar.AddForm NewCodeWindow
            Else
                frmMain.TabBar.SwitchToByForm CurrentProject.Files(TvItemBinding(i).FileIndex).TargetWindow '切换到对应的窗口
                CurrentProject.Files(TvItemBinding(i).FileIndex).TargetWindow.SyntaxEdit.SetFocus
            End If
            Exit For
        End If
    Next i
End Sub

Public Sub SolutionTreeView_EndLabelEdit(ByVal hTreeItem As Long, NewText As String, bCancel As Boolean)
    If NewText = vbNullChar Then                                                                '如果NewText为vbNullChar，则说明编辑被取消了
        Exit Sub
    Else                                                                                        '尝试进行重命名
        On Error Resume Next
        Dim i   As Long
        
        If hTreeItem = ProjectNameTvItem Then                                                       '如果列表项对应的是工程名称，则更改工程文件名
            Name ProjectFilePath As ProjectFolderPath & NewText & ".myproj"
            If Err.Number <> 0 Then                                                                     '重命名时发生错误
                NoSkinMsgBox Lang_SolutionExplorer_RenameFailure_1 & ProjectFilePath & _
                    Lang_SolutionExplorer_RenameFailure_2 & Err.Number & " - " & Err.Description, vbExclamation, Lang_Msgbox_Error
                bCancel = True
            Else                                                                                        '重命名成功
                ProjectFilePath = ProjectFolderPath & NewText & ".myproj"                                   '更新工程文件路径
                CurrentProject.ProjectName = NewText                                                        '更新工程名称
                CurrentProject.Changed = True                                                               '标记工程已更改
            End If
            Exit Sub
        End If
        For i = 0 To UBound(TvItemBinding)                                                          '查找列表项对应的文件序号
            If hTreeItem = TvItemBinding(i).TVITEM Then                                                 '找到匹配的文件就进行重命名
                Name CurrentProject.Files(i).FilePath As ProjectFolderPath & NewText
                If Err.Number <> 0 Then                                                                     '重命名时发生错误
                    NoSkinMsgBox Lang_SolutionExplorer_RenameFailure_1 & CurrentProject.Files(i).FilePath & _
                        Lang_SolutionExplorer_RenameFailure_2 & Err.Number & " - " & Err.Description, vbExclamation, Lang_Msgbox_Error
                    bCancel = True
                Else                                                                                        '重命名成功
                    CurrentProject.Files(i).TargetWindow.Caption = NewText                                      '刷新窗口标题
                    frmMain.TabBar.UpdateCaptions
                    CurrentProject.Files(i).FilePath = ProjectFolderPath & NewText                              '更新文件路径
                End If
                Exit Sub
            End If
        Next i
    End If
    bCancel = True                                                                              '其实应该不会找不到对应的文件，但是如果真的找不到就取消操作吧
End Sub

Public Sub SolutionTreeView_ItemExpanding(ByVal hTreeItem As Long, bCancel As Boolean)

End Sub

Public Sub SolutionTreeView_KeyDown(ByVal KeyCode As Long, ByVal IsLongPress As Boolean)
    If KeyCode = vbKeyF2 Then                                                                   '响应F2键: 重命名
        Me.SolutionTreeView.EditLabel Me.SolutionTreeView.GetSelectedItem()
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

Public Sub SolutionTreeView_RightClick(bCancel As Boolean)
    Dim i               As Long
    Dim hTreeItem       As Long
    
    '判断选定的列表项的类型并弹出对应的菜单
    hTreeItem = Me.SolutionTreeView.GetSelectedItem()
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
        Me.mnuProjectItemPopup.PopupMenu 0
    Else                                                                                        '否则检查选择的项目是否为代码文件
        For i = 0 To UBound(TvItemBinding)
            If hTreeItem = TvItemBinding(i).TVITEM Then                                             '如果能找到对应的文件，说明选择的列表项是文件而不是项目节点
                Me.mnuItemPopup.PopupMenu 0
                
                Exit Sub
            End If
        Next i
        Me.mnuProjectItemPopup.PopupMenu 0                                                      '不能找到对应的文件说明是项目节点
    End If
End Sub

Public Sub SolutionTreeView_SelChanged(ByVal hOldTreeItem As Long, ByVal hNewTreeItem As Long)
    
End Sub

Public Sub SolutionTreeView_SelChanging(ByVal hOldTreeItem As Long, ByVal hNewTreeItem As Long, bCancel As Boolean)
    
End Sub

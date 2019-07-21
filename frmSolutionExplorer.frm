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
   Begin DragControlsIDE.DarkTreeView SolutionTreeView 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _extentx        =   5318
      _extenty        =   5106
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
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.SolutionTreeView.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
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
                Dim FileTitle       As String                                                               '文件名
                
                Set NewCodeWindow = CreateNewCodeWindow(TvItemBinding(i).FileIndex)                         '创建新的代码窗体并设置绑定的文件序号
                FileTitle = CurrentProject.Files(TvItemBinding(i).FileIndex).FilePath
                FileTitle = Right(FileTitle, Len(FileTitle) - InStrRev(FileTitle, "\"))                     '截取出文件名
                NewCodeWindow.Caption = FileTitle
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
    If KeyCode = vbKeyF2 Then                                                                   '响应F2键：重命名
        Me.SolutionTreeView.EditLabel Me.SolutionTreeView.GetSelectedItem()
    End If
End Sub

Public Sub SolutionTreeView_KeyUp(ByVal KeyCode As Long)

End Sub

Public Sub SolutionTreeView_MouseDown(ByVal Button As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Public Sub SolutionTreeView_MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Public Sub SolutionTreeView_MouseUp(ByVal Button As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Public Sub SolutionTreeView_RightClick(bCancel As Boolean)

End Sub

Public Sub SolutionTreeView_SelChanged(ByVal hOldTreeItem As Long, ByVal hNewTreeItem As Long)
    
End Sub

Public Sub SolutionTreeView_SelChanging(ByVal hOldTreeItem As Long, ByVal hNewTreeItem As Long, bCancel As Boolean)
    
End Sub

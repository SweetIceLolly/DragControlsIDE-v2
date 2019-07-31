VERSION 5.00
Begin VB.Form frmSolutionExplorer 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "������Դ������"
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
'����:      ������Դ��������������ʾ������������Ŀ¼���ļ�
'����:      ����
'�ļ�:      frmSolutionExplorer.frm
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
    
    If hTreeItem = ProjectNameTvItem Then                                                       '����б����Ӧ���ǹ������ƣ���Ҳ�������
        Exit Sub
    End If
    For i = 0 To UBound(TvItemBinding)                                                          '�����б����Ӧ���ļ����
        If hTreeItem = TvItemBinding(i).TVITEM Then                                                 '������ҵ���Ӧ���ļ���˵��ѡ����б������ļ���������Ŀ�ڵ�
            '�����ǩ�����С�.������ôֻѡ��.��ǰ����ı�
            Dim hwndLabelEditBox    As Long                                                             '���б�ǩ�༭���ı�����
            Dim LabelStr            As String                                                           '��ǰ׼���༭�ı�ǩ���ı�
            Dim DotPos              As Integer                                                          'С���㡰.���ڱ�ǩ�ı����λ��
            
            LabelStr = Me.SolutionTreeView.GetItemText(hTreeItem)                                       '��ȡ��ǰ׼���༭�ı�ǩ���ı�
            DotPos = InStrRev(LabelStr, ".")                                                            '���ı��в��ҡ�.��
            If DotPos <> 0 Then                                                                         '����ҵ�С����
                hwndLabelEditBox = SendMessageA(Me.SolutionTreeView.TreeViewHwnd, TVM_GETEDITCONTROL, 0, 0) '��ȡ���б�ǩ�༭���ı�����
                SetPropA hwndLabelEditBox, "PrevWndProc", _
                    SetWindowLongA(hwndLabelEditBox, GWL_WNDPROC, AddressOf TreeViewEditBoxWindowProc)      '���ñ�ǩ�༭���ı�������໯������ѡ���ı�����Ϣ
                SetPropA hwndLabelEditBox, "DotPos", ByVal DotPos - 1                                       '��¼��.����λ�ã��Ա��ı�������໯�޸�ѡ����ı�
            End If
            
            Exit Sub
        End If
    Next i
    bCancel = True                                                                              '����Ҳ�����Ӧ���ļ���˵��ѡ����б����ǲ���������������Ŀ�ڵ�
End Sub

Public Sub SolutionTreeView_Click(bCancel As Boolean)
    
End Sub

Public Sub SolutionTreeView_DoubleClick(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    On Error Resume Next
    Dim CurrSelItem     As Long
    Dim i               As Long
    
    CurrSelItem = Me.SolutionTreeView.GetSelectedItem()                                         '��ȡѡ�������ͼ�б���
    For i = 0 To UBound(TvItemBinding)                                                          '�����б����Ӧ���ļ����
        If CurrSelItem = TvItemBinding(i).TVITEM Then
            If CurrentProject.Files(TvItemBinding(i).FileIndex).TargetWindow Is Nothing Then
                Dim NewCodeWindow   As frmCodeWindow                                                        '�½��Ĵ������
                
                Set NewCodeWindow = CreateNewCodeWindow(TvItemBinding(i).FileIndex)                         '�����µĴ��봰�岢���ð󶨵��ļ����
                NewCodeWindow.Caption = GetFileName(CurrentProject.Files(TvItemBinding(i).FileIndex).FilePath)
                frmMain.TabBar.AddForm NewCodeWindow
            Else
                frmMain.TabBar.SwitchToByForm CurrentProject.Files(TvItemBinding(i).FileIndex).TargetWindow '�л�����Ӧ�Ĵ���
                CurrentProject.Files(TvItemBinding(i).FileIndex).TargetWindow.SyntaxEdit.SetFocus
            End If
            Exit For
        End If
    Next i
End Sub

Public Sub SolutionTreeView_EndLabelEdit(ByVal hTreeItem As Long, NewText As String, bCancel As Boolean)
    If NewText = vbNullChar Then                                                                '���NewTextΪvbNullChar����˵���༭��ȡ����
        Exit Sub
    Else                                                                                        '���Խ���������
        On Error Resume Next
        Dim i   As Long
        
        If hTreeItem = ProjectNameTvItem Then                                                       '����б����Ӧ���ǹ������ƣ�����Ĺ����ļ���
            Name ProjectFilePath As ProjectFolderPath & NewText & ".myproj"
            If Err.Number <> 0 Then                                                                     '������ʱ��������
                NoSkinMsgBox Lang_SolutionExplorer_RenameFailure_1 & ProjectFilePath & _
                    Lang_SolutionExplorer_RenameFailure_2 & Err.Number & " - " & Err.Description, vbExclamation, Lang_Msgbox_Error
                bCancel = True
            Else                                                                                        '�������ɹ�
                ProjectFilePath = ProjectFolderPath & NewText & ".myproj"                                   '���¹����ļ�·��
                CurrentProject.ProjectName = NewText                                                        '���¹�������
                CurrentProject.Changed = True                                                               '��ǹ����Ѹ���
            End If
            Exit Sub
        End If
        For i = 0 To UBound(TvItemBinding)                                                          '�����б����Ӧ���ļ����
            If hTreeItem = TvItemBinding(i).TVITEM Then                                                 '�ҵ�ƥ����ļ��ͽ���������
                Name CurrentProject.Files(i).FilePath As ProjectFolderPath & NewText
                If Err.Number <> 0 Then                                                                     '������ʱ��������
                    NoSkinMsgBox Lang_SolutionExplorer_RenameFailure_1 & CurrentProject.Files(i).FilePath & _
                        Lang_SolutionExplorer_RenameFailure_2 & Err.Number & " - " & Err.Description, vbExclamation, Lang_Msgbox_Error
                    bCancel = True
                Else                                                                                        '�������ɹ�
                    CurrentProject.Files(i).TargetWindow.Caption = NewText                                      'ˢ�´��ڱ���
                    frmMain.TabBar.UpdateCaptions
                    CurrentProject.Files(i).FilePath = ProjectFolderPath & NewText                              '�����ļ�·��
                End If
                Exit Sub
            End If
        Next i
    End If
    bCancel = True                                                                              '��ʵӦ�ò����Ҳ�����Ӧ���ļ��������������Ҳ�����ȡ��������
End Sub

Public Sub SolutionTreeView_ItemExpanding(ByVal hTreeItem As Long, bCancel As Boolean)

End Sub

Public Sub SolutionTreeView_KeyDown(ByVal KeyCode As Long, ByVal IsLongPress As Boolean)
    If KeyCode = vbKeyF2 Then                                                                   '��ӦF2����������
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

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
      SubMenuText_1_1 =   "��(&O)"
      SubMenuID_1_1   =   2
      SubMenuText_1_2 =   "������(&R)"
      SubMenuID_1_2   =   3
      SubMenuText_1_3 =   "����(&C)"
      SubMenuID_1_3   =   4
      SubMenuText_1_4 =   "����Ŀ�Ƴ�(&E)"
      SubMenuID_1_4   =   5
      SubMenuText_1_5 =   "ɾ��(&D)"
      SubMenuID_1_5   =   6
      SubMenuText_1_6 =   "���ļ�������д�(&P)"
      SubMenuID_1_6   =   7
      MenuID_2        =   1
      MenuText_2      =   "��(&O)"
      MenuVisible_2   =   -1  'True
      MenuIcon_2      =   "frmSolutionExplorer.frx":0018
      SubMenuID_2_0   =   0
      MenuID_3        =   2
      MenuText_3      =   "������(&R)"
      MenuVisible_3   =   -1  'True
      MenuIcon_3      =   "frmSolutionExplorer.frx":0030
      SubMenuID_3_0   =   0
      MenuID_4        =   3
      MenuText_4      =   "����(&C)"
      MenuVisible_4   =   -1  'True
      MenuIcon_4      =   "frmSolutionExplorer.frx":0048
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "����Ŀ�Ƴ�(&E)"
      MenuVisible_5   =   -1  'True
      MenuIcon_5      =   "frmSolutionExplorer.frx":0060
      SubMenuID_5_0   =   0
      MenuID_6        =   5
      MenuText_6      =   "ɾ��(&D)"
      MenuVisible_6   =   -1  'True
      MenuIcon_6      =   "frmSolutionExplorer.frx":0078
      SubMenuID_6_0   =   0
      MenuID_7        =   6
      MenuText_7      =   "���ļ�������д�(&P)"
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
      SubMenuText_1_1 =   "���빤��(&C)"
      SubMenuID_1_1   =   2
      SubMenuText_1_2 =   "���"
      SubMenuID_1_2   =   3
      SubMenuText_1_3 =   "������(&R)"
      SubMenuID_1_3   =   11
      SubMenuText_1_4 =   "���ļ�������д�(&O)"
      SubMenuID_1_4   =   12
      SubMenuText_1_5 =   "��������(&P)"
      SubMenuID_1_5   =   13
      MenuID_2        =   1
      MenuText_2      =   "���빤��(&C)"
      MenuVisible_2   =   -1  'True
      MenuIcon_2      =   "frmSolutionExplorer.frx":00C0
      SubMenuID_2_0   =   0
      MenuID_3        =   2
      MenuText_3      =   "���"
      MenuVisible_3   =   -1  'True
      MenuIcon_3      =   "frmSolutionExplorer.frx":00D8
      SUBMENU_ITEM_COUNT_3=   7
      SubMenuID_3_0   =   0
      SubMenuText_3_1 =   "�ļ���(&F)"
      SubMenuID_3_1   =   4
      SubMenuText_3_2 =   "-"
      SubMenuID_3_2   =   5
      SubMenuText_3_3 =   "����(&W)"
      SubMenuID_3_3   =   6
      SubMenuText_3_4 =   "C++�ļ� (.cpp)"
      SubMenuID_3_4   =   7
      SubMenuText_3_5 =   "C++ͷ�ļ� (.hpp)"
      SubMenuID_3_5   =   8
      SubMenuText_3_6 =   "C�ļ� (.c)"
      SubMenuID_3_6   =   9
      SubMenuText_3_7 =   "Cͷ�ļ� (.h)"
      SubMenuID_3_7   =   10
      MenuID_4        =   3
      MenuText_4      =   "�ļ���(&F)"
      MenuVisible_4   =   -1  'True
      MenuIcon_4      =   "frmSolutionExplorer.frx":00F0
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "-"
      MenuVisible_5   =   -1  'True
      MenuIcon_5      =   "frmSolutionExplorer.frx":0108
      SubMenuID_5_0   =   0
      MenuID_6        =   5
      MenuText_6      =   "����(&W)"
      MenuVisible_6   =   -1  'True
      MenuIcon_6      =   "frmSolutionExplorer.frx":0120
      SubMenuID_6_0   =   0
      MenuID_7        =   6
      MenuText_7      =   "C++�ļ� (.cpp)"
      MenuVisible_7   =   -1  'True
      MenuIcon_7      =   "frmSolutionExplorer.frx":0138
      SubMenuID_7_0   =   0
      MenuID_8        =   7
      MenuText_8      =   "C++ͷ�ļ� (.hpp)"
      MenuVisible_8   =   -1  'True
      MenuIcon_8      =   "frmSolutionExplorer.frx":0150
      SubMenuID_8_0   =   0
      MenuID_9        =   8
      MenuText_9      =   "C�ļ� (.c)"
      MenuVisible_9   =   -1  'True
      MenuIcon_9      =   "frmSolutionExplorer.frx":0168
      SubMenuID_9_0   =   0
      MenuID_10       =   9
      MenuText_10     =   "Cͷ�ļ� (.h)"
      MenuVisible_10  =   -1  'True
      MenuIcon_10     =   "frmSolutionExplorer.frx":0180
      SubMenuID_10_0  =   0
      MenuID_11       =   10
      MenuText_11     =   "������(&R)"
      MenuVisible_11  =   -1  'True
      MenuIcon_11     =   "frmSolutionExplorer.frx":0198
      SubMenuID_11_0  =   0
      MenuID_12       =   11
      MenuText_12     =   "���ļ�������д�(&O)"
      MenuVisible_12  =   -1  'True
      MenuIcon_12     =   "frmSolutionExplorer.frx":01B0
      SubMenuID_12_0  =   0
      MenuID_13       =   12
      MenuText_13     =   "��������(&P)"
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
'����:      ������Դ��������������ʾ������������Ŀ¼���ļ�
'����:      ����
'�ļ�:      frmSolutionExplorer.frm
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
        Case 1                                  '��
            Call SolutionTreeView_DoubleClick(1, 0, 0, 0)
        
        Case 2                                  '������
            Me.SolutionTreeView.EditLabel Me.SolutionTreeView.GetSelectedItem()
        
        Case 3                                  '����
        
        Case 4                                  '����Ŀ�Ƴ�
        
        Case 5                                  'ɾ��
        
        Case 6                                  '���ļ��������
            Dim i           As Long
            Dim hTreeItem    As Long
            
            hTreeItem = Me.SolutionTreeView.GetSelectedItem()
            For i = 0 To UBound(TvItemBinding)                              '�����б����Ӧ���ļ����
                If hTreeItem = TvItemBinding(i).TVITEM Then                     '�ҵ���Ӧ���ļ�
                    Shell "explorer.exe /select,""" & CurrentProject.Files(TvItemBinding(i).FileIndex).FilePath & """", vbNormalFocus
                End If
            Next i
        
    End Select
End Sub

Private Sub mnuProjectItemPopup_MenuItemClicked(MenuID As Integer)
    Me.mnuProjectItemPopup.HideMenu
    Select Case MenuID
        Case 1                                  '���빤��
            
        
        Case 3                                  '�ļ���
        
        Case 5                                  '��Ӵ���
            
        Case 6                                  '���cpp
        
        Case 7                                  '���hpp
        
        Case 8                                  '���c
        
        Case 9                                  '���h
        
        Case 10                                 '������
            Me.SolutionTreeView.EditLabel Me.SolutionTreeView.GetSelectedItem()
        
        Case 11                                 '���ļ��������
            Shell "explorer.exe /select,""" & ProjectFilePath & """", vbNormalFocus
        
        Case 12                                 '��������
        
    End Select
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
    If KeyCode = vbKeyF2 Then                                                                   '��ӦF2��: ������
        Me.SolutionTreeView.EditLabel Me.SolutionTreeView.GetSelectedItem()
    End If
End Sub

Public Sub SolutionTreeView_KeyUp(ByVal KeyCode As Long)

End Sub

Public Sub SolutionTreeView_MouseDown(ByVal Button As Long, ByVal X As Long, ByVal Y As Long)
    Me.SolutionTreeView.SelectItem Me.SolutionTreeView.HitTest(X, Y)                            'ѡ����갴�µ�λ�õ��б���
End Sub

Public Sub SolutionTreeView_MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Public Sub SolutionTreeView_MouseUp(ByVal Button As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Public Sub SolutionTreeView_RightClick(bCancel As Boolean)
    Dim i               As Long
    Dim hTreeItem       As Long
    
    '�ж�ѡ�����б�������Ͳ�������Ӧ�Ĳ˵�
    hTreeItem = Me.SolutionTreeView.GetSelectedItem()
    If hTreeItem = 0 Then
        Exit Sub
    End If
    If CurrentProject.ProjectType = 1 Then                                                      '����������Ͳ��Ǵ��ڳ���Ͳ�������Ӵ���
        Me.mnuProjectItemPopup.MenuEnabled(5) = True
    Else
        Me.mnuProjectItemPopup.MenuEnabled(5) = False
    End If
    frmMain.DarkTitleBar.SetFocus
    If hTreeItem = ProjectNameTvItem Then                                                       '���ѡ�����Ŀ�ǹ����ļ�
        Me.mnuProjectItemPopup.PopupMenu 0
    Else                                                                                        '������ѡ�����Ŀ�Ƿ�Ϊ�����ļ�
        For i = 0 To UBound(TvItemBinding)
            If hTreeItem = TvItemBinding(i).TVITEM Then                                             '������ҵ���Ӧ���ļ���˵��ѡ����б������ļ���������Ŀ�ڵ�
                Me.mnuItemPopup.PopupMenu 0
                
                Exit Sub
            End If
        Next i
        Me.mnuProjectItemPopup.PopupMenu 0                                                      '�����ҵ���Ӧ���ļ�˵������Ŀ�ڵ�
    End If
End Sub

Public Sub SolutionTreeView_SelChanged(ByVal hOldTreeItem As Long, ByVal hNewTreeItem As Long)
    
End Sub

Public Sub SolutionTreeView_SelChanging(ByVal hOldTreeItem As Long, ByVal hNewTreeItem As Long, bCancel As Boolean)
    
End Sub

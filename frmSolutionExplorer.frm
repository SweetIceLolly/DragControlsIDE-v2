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
      SubMenuText_1_1 =   "��(&O)"
      SubMenuID_1_1   =   2
      SubMenuText_1_2 =   "������(&R)"
      SubMenuID_1_2   =   3
      SubMenuText_1_3 =   "����Ŀ�Ƴ�(&E)"
      SubMenuID_1_3   =   4
      SubMenuText_1_4 =   "���ļ�������д�(&P)"
      SubMenuID_1_4   =   5
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
      MenuText_4      =   "����Ŀ�Ƴ�(&E)"
      MenuVisible_4   =   -1  'True
      MenuIcon_4      =   "frmSolutionExplorer.frx":0048
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "���ļ�������д�(&P)"
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
      SubMenuText_1_1 =   "���빤��(&C)"
      SubMenuID_1_1   =   2
      SubMenuText_1_2 =   "���"
      SubMenuID_1_2   =   3
      SubMenuText_1_3 =   "������(&R)"
      SubMenuID_1_3   =   11
      SubMenuText_1_4 =   "ɾ��(&D)"
      SubMenuID_1_4   =   12
      SubMenuText_1_5 =   "���ļ�������д�(&O)"
      SubMenuID_1_5   =   13
      SubMenuText_1_6 =   "��������(&P)"
      SubMenuID_1_6   =   14
      MenuID_2        =   1
      MenuText_2      =   "���빤��(&C)"
      MenuVisible_2   =   -1  'True
      MenuIcon_2      =   "frmSolutionExplorer.frx":0090
      SubMenuID_2_0   =   0
      MenuID_3        =   2
      MenuText_3      =   "���"
      MenuVisible_3   =   -1  'True
      MenuIcon_3      =   "frmSolutionExplorer.frx":00A8
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
      MenuIcon_4      =   "frmSolutionExplorer.frx":00C0
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "-"
      MenuVisible_5   =   -1  'True
      MenuIcon_5      =   "frmSolutionExplorer.frx":00D8
      SubMenuID_5_0   =   0
      MenuID_6        =   5
      MenuText_6      =   "����(&W)"
      MenuVisible_6   =   -1  'True
      MenuIcon_6      =   "frmSolutionExplorer.frx":00F0
      SubMenuID_6_0   =   0
      MenuID_7        =   6
      MenuText_7      =   "C++�ļ� (.cpp)"
      MenuVisible_7   =   -1  'True
      MenuIcon_7      =   "frmSolutionExplorer.frx":0108
      SubMenuID_7_0   =   0
      MenuID_8        =   7
      MenuText_8      =   "C++ͷ�ļ� (.hpp)"
      MenuVisible_8   =   -1  'True
      MenuIcon_8      =   "frmSolutionExplorer.frx":0120
      SubMenuID_8_0   =   0
      MenuID_9        =   8
      MenuText_9      =   "C�ļ� (.c)"
      MenuVisible_9   =   -1  'True
      MenuIcon_9      =   "frmSolutionExplorer.frx":0138
      SubMenuID_9_0   =   0
      MenuID_10       =   9
      MenuText_10     =   "Cͷ�ļ� (.h)"
      MenuVisible_10  =   -1  'True
      MenuIcon_10     =   "frmSolutionExplorer.frx":0150
      SubMenuID_10_0  =   0
      MenuID_11       =   10
      MenuText_11     =   "������(&R)"
      MenuVisible_11  =   -1  'True
      MenuIcon_11     =   "frmSolutionExplorer.frx":0168
      SubMenuID_11_0  =   0
      MenuID_12       =   11
      MenuText_12     =   "ɾ��(&D)"
      MenuVisible_12  =   -1  'True
      MenuIcon_12     =   "frmSolutionExplorer.frx":0180
      SubMenuID_12_0  =   0
      MenuID_13       =   12
      MenuText_13     =   "���ļ�������д�(&O)"
      MenuVisible_13  =   -1  'True
      MenuIcon_13     =   "frmSolutionExplorer.frx":0198
      SubMenuID_13_0  =   0
      MenuID_14       =   13
      MenuText_14     =   "��������(&P)"
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
'����:      ������Դ��������������ʾ������������Ŀ¼���ļ�
'����:      ����
'�ļ�:      frmSolutionExplorer.frm
'====================================================

Option Explicit

'���������������ڴ����ļ���
Dim IsCreatingFolder    As Boolean                                  '�Ƿ����ڴ����ļ���
Dim IsCreatingFile      As Boolean                                  '�Ƿ����ڴ����ļ�
Dim CreatedTreeItem     As Long                                     '���ڴ������ļ��л����ļ�������ͼ�ڵ�
Dim ParentOfCreating    As Long                                     '���ڴ������ļ��л����ļ�������ͼ�ڵ��ĸ�ڵ�
Dim CreatingDefaultName As String                                   '���ڴ������ļ��л����ļ���Ĭ������

'����:      �ݹ�����ļ����µ��ӽڵ�·��
'����:      ParentIndex: ĸ�ļ������
Private Sub RenameFolder(ParentIndex As Long)
    Dim i               As Long
    
    For i = 0 To UBound(CurrentProject.Folders)                     '��������ļ��У������ĸ�ļ��б�����������ô�͸�������·��
        If CurrentProject.Folders(i).ParentFolder = ParentIndex Then
            CurrentProject.Folders(i).FolderPath = CurrentProject.Folders(ParentIndex).FolderPath & "\" & _
                GetFileName(CurrentProject.Folders(i).FolderPath)
            RenameFolder i                                                  '������һ���ļ��е�·��
        End If
    Next i
    For i = 0 To UBound(CurrentProject.Files)                       '��������ļ���������ļ��б�����������ô�͸�������·��
        If CurrentProject.Files(i).FolderIndex = ParentIndex Then
            CurrentProject.Files(i).FilePath = ProjectFolderPath & CurrentProject.Folders(ParentIndex).FolderPath & "\" & _
                GetFileName(CurrentProject.Files(i).FilePath)
        End If
    Next i
End Sub

'����:      �����ļ�������򿪡��˵�
Private Sub mnuOpenWithExplorer_Click()
    Dim i               As Long
    Dim hTreeItem       As Long
    
    hTreeItem = Me.SolutionTreeView.GetSelectedItem()
    If hTreeItem = ProjectNameTvItem Then                           'ѡ����б�������Ŀ�ڵ�
        Shell "explorer.exe /select,""" & ProjectFilePath & """", vbNormalFocus
        Exit Sub
    End If
    For i = 0 To UBound(TvItemBinding)                              '�����б����Ӧ���ļ����
        If hTreeItem = TvItemBinding(i).TVITEM Then                     '�ҵ���Ӧ���ļ�
            If TvItemBinding(i).IsFolder Then                               'ѡ�����Ŀ���ļ���
                Shell "explorer.exe """ & ProjectFolderPath & CurrentProject.Folders(TvItemBinding(i).FileIndex).FolderPath & """", vbNormalFocus
            Else                                                            'ѡ�����Ŀ���ļ�
                Shell "explorer.exe /select,""" & CurrentProject.Files(TvItemBinding(i).FileIndex).FilePath & """", vbNormalFocus
            End If
        End If
    Next i
End Sub

'����:      ���½��ļ��С��˵�
Private Sub mnuCreateFolder_Click()
    Dim hTreeItem       As Long
    
    IsCreatingFolder = True                                                         '���Ϊ���ڴ����ļ���
    IsCreatingFile = False
    CreatingDefaultName = Lang_SolutionExplorer_NewFolderName
    hTreeItem = Me.SolutionTreeView.GetSelectedItem()
    CreatedTreeItem = Me.SolutionTreeView.AddItem(CreatingDefaultName, hTreeItem)   '�����ļ��нڵ�
    Me.SolutionTreeView.ExpandItems hTreeItem, 2
    ParentOfCreating = hTreeItem
    Me.SolutionTreeView.EditLabel CreatedTreeItem                                   '��ʼ�༭��ǩ
End Sub

'����:      ����ļ�����
'����:      FileName: ��ӵ��ļ���
Private Sub mnuAddFile_Click(FileName As String)
    Dim hTreeItem       As Long
    
    IsCreatingFile = True                                                           '���Ϊ���ڴ����ļ�
    IsCreatingFolder = False
    CreatingDefaultName = FileName                                                  '����Ĭ������
    hTreeItem = Me.SolutionTreeView.GetSelectedItem()
    CreatedTreeItem = Me.SolutionTreeView.AddItem(FileName, hTreeItem)              '�����ļ��ڵ�
    Me.SolutionTreeView.ExpandItems hTreeItem, 2
    ParentOfCreating = hTreeItem
    Me.SolutionTreeView.EditLabel CreatedTreeItem                                   '��ʼ�༭��ǩ
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
        Case 1                                  '��
            Call SolutionTreeView_DoubleClick(1, 0, 0, 0)
        
        Case 2                                  '������
            Me.SolutionTreeView.EditLabel Me.SolutionTreeView.GetSelectedItem()
        
        Case 3                                  '����Ŀ�Ƴ�
        
        Case 4                                  '���ļ��������
            Call mnuOpenWithExplorer_Click
        
    End Select
End Sub

Private Sub mnuProjectItemPopup_MenuItemClicked(MenuID As Integer)
    Me.mnuProjectItemPopup.HideMenu
    Select Case MenuID
        Case 1                                  '���빤��
        
        Case 3                                  '�ļ���
            Call mnuCreateFolder_Click
        
        Case 5                                  '��Ӵ���
            
        Case 6                                  '���cpp
            Call mnuAddFile_Click("��cpp�ļ�.cpp")
        
        Case 7                                  '���hpp
            Call mnuAddFile_Click("��hpp�ļ�.hpp")
        
        Case 8                                  '���c
            Call mnuAddFile_Click("��c�ļ�.c")
        
        Case 9                                  '���h
            Call mnuAddFile_Click("��h�ļ�.h")
        
        Case 10                                 '������
            Me.SolutionTreeView.EditLabel Me.SolutionTreeView.GetSelectedItem()
        
        Case 11                                 'ɾ��
            
        
        Case 12                                 '���ļ��������
            Call mnuOpenWithExplorer_Click
        
        Case 13                                 '��������
        
    End Select
End Sub

Public Sub SolutionTreeView_BeginLabelEdit(ByVal hTreeItem As Long, bCancel As Boolean)
    Dim i               As Long
    
    If IsCreatingFolder Then                                                                    '������ڴ����ļ��У����������
        Exit Sub
    End If
    If IsCreatingFile Then                                                                      '������ڴ����ļ������Զ�ѡȡС����ǰ����ı�
        GoTo SelectEditboxText
    End If
    
    If hTreeItem = ProjectNameTvItem Then                                                       '����б����Ӧ���ǹ������ƣ����������
        Exit Sub
    End If
    For i = 0 To UBound(TvItemBinding)                                                          '�����б����Ӧ���ļ����
        If hTreeItem = TvItemBinding(i).TVITEM Then                                                 '������ҵ���Ӧ���ļ���˵��ѡ����б������ļ���������Ŀ�ڵ�
            If TvItemBinding(i).IsFolder Then                                                           'ѡ�����Ŀ���ļ��У������޸�
                Exit Sub
            Else                                                                                        'ѡ�����Ŀ���ļ�
                GoTo SelectEditboxText
            End If
            Exit Sub
        End If
    Next i
    bCancel = True                                                                              '����Ҳ�����Ӧ���ļ���˵��ѡ����б����ǲ���������������Ŀ�ڵ�
    
SelectEditboxText:
    Dim hwndLabelEditBox    As Long                                                             '���б�ǩ�༭���ı�����
    Dim LabelStr            As String                                                           '��ǰ׼���༭�ı�ǩ���ı�
    Dim DotPos              As Integer                                                          'С���㡰.���ڱ�ǩ�ı����λ��
    
    '�����ǩ�����С�.������ôֻѡ��.��ǰ����ı�
    LabelStr = Me.SolutionTreeView.GetItemText(hTreeItem)                                       '��ȡ��ǰ׼���༭�ı�ǩ���ı�
    DotPos = InStrRev(LabelStr, ".")                                                            '���ı��в��ҡ�.��
    If DotPos <> 0 Then                                                                         '����ҵ�С����
        hwndLabelEditBox = SendMessageA(Me.SolutionTreeView.TreeViewHwnd, TVM_GETEDITCONTROL, 0, 0) '��ȡ���б�ǩ�༭���ı�����
        SetPropA hwndLabelEditBox, "PrevWndProc", _
            SetWindowLongA(hwndLabelEditBox, GWL_WNDPROC, AddressOf TreeViewEditBoxWindowProc)      '���ñ�ǩ�༭���ı�������໯������ѡ���ı�����Ϣ
        SetPropA hwndLabelEditBox, "DotPos", ByVal DotPos - 1                                       '��¼��.����λ�ã��Ա��ı�������໯�޸�ѡ����ı�
    End If
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
            If TvItemBinding(i).IsFolder Then                                                           '���ѡ�����Ŀ���ļ���
                Me.SolutionTreeView.ExpandItems CurrSelItem, 3                                              '�л��ڵ�չ��״̬
                Me.SolutionTreeView.EndEditLabel False                                                      'ȡ���༭��ǩ
            Else                                                                                        '���ѡ�����Ŀ�Ǵ����ļ�
                Dim NewCodeWindow   As frmCodeWindow                                                        '��Ӧ�Ĵ������
                
                Set NewCodeWindow = frmMain.ShowCodeWindow(TvItemBinding(i).FileIndex)                      '��ȡ��Ŀ�ڵ�����Ӧ�Ĵ��봰��
                If NewCodeWindow Is Nothing Then
                    NoSkinMsgBox Lang_Main_Debug_OpenSourceFailure & CurrentProject.Files(TvItemBinding(i).FileIndex).FilePath, _
                        vbExclamation, Lang_Msgbox_Error
                Else
                    NewCodeWindow.SyntaxEdit.SetFocus
                End If
            End If
            Exit For
        End If
    Next i
End Sub

Public Sub SolutionTreeView_EndLabelEdit(ByVal hTreeItem As Long, NewText As String, bCancel As Boolean)
    If NewText = vbNullChar Or NewText = "" Then                                                '���NewTextΪvbNullChar����˵���༭��ȡ����
        If IsCreatingFolder Or IsCreatingFile Then                                                  '������ڴ����ļ��л����ļ���ʹ��Ĭ������
            NewText = CreatingDefaultName
        Else                                                                                        '���Ǵ����ļ��еĻ���ȡ��������
            Exit Sub
        End If
    End If
    
    If NewText = "" Then                                                                        '�ǿ��ı�
        If IsCreatingFolder Or IsCreatingFile Then                                                  '������ڴ����ļ������ļ��о�ȡ������
            IsCreatingFolder = False
            IsCreatingFile = False
            Me.SolutionTreeView.RemoveItem CreatedTreeItem
        End If
        Exit Sub
    End If
    
    '���Խ���������
    On Error Resume Next
    Dim i   As Long
    
    If Not CheckInvalidFileName(NewText) Then                                                   '���Ƿ��ļ���
        If IsCreatingFolder Or IsCreatingFile Then                                                  '������ڴ����ļ������ļ��о�ȡ������
            IsCreatingFolder = False
            IsCreatingFile = False
            Me.SolutionTreeView.RemoveItem CreatedTreeItem
        End If
        NoSkinMsgBox Lang_SolutionExplorer_InvalidName, vbExclamation, Lang_Msgbox_Error
        bCancel = True
        Exit Sub
    End If
    
    If IsCreatingFolder Or IsCreatingFile Then                                                      '������ڴ����ļ������ļ���
        Dim FolderPath          As String                                                           '�����ļ������ļ��е�λ��
        Dim ParentFolderIndex   As Long                                                             '�������ļ������ļ��еĽڵ��ĸ�ڵ������
        
        For i = 0 To UBound(TvItemBinding)                                                          '���Ҷ�Ӧ��ĸ�ڵ���CurrentProject.Folders������
            If ParentOfCreating = TvItemBinding(i).TVITEM Then                                          '��¼��ĸ�ڵ���ƥ�����������ȡĸ�ڵ�·�����γ����������·��
                ParentFolderIndex = TvItemBinding(i).FileIndex
                FolderPath = CurrentProject.Folders(ParentFolderIndex).FolderPath & "\" & FolderPath
                Exit For
            End If
        Next i
        
        Err.Clear
        If IsCreatingFolder Then                                                                    '����Ǵ����ļ���
            MkDir ProjectFolderPath & FolderPath & NewText
        Else
            If Dir(ProjectFolderPath & FolderPath & NewText, vbDirectory Or vbNormal Or vbReadOnly Or vbHidden Or vbSystem) = "" Then
                Open ProjectFolderPath & FolderPath & NewText For Binary As #1
                Close #1
            Else                                                                                    '��⵽����
                Err.Raise 75
            End If
        End If
        If Err.Number <> 0 Then                                                                     '�����ļ���ʱ��������
            Me.SolutionTreeView.RemoveItem CreatedTreeItem
            MsgBox "����ʧ�ܣ�", vbExclamation, Lang_Msgbox_Error       'todo
        Else                                                                                        '�����ļ������ļ��гɹ�
            CurrentProject.Changed = True                                                               '��ǹ����ļ�Ϊ�Ѹ���
            If IsCreatingFolder Then                                                                    '����Ǵ����ļ��о͸�����Ŀ��Ϣ����ļ�����Ϣ
                ReDim Preserve CurrentProject.Folders(UBound(CurrentProject.Folders) + 1)                   '�����Ŀ��Ϣ����ļ�����Ϣ
                ReDim Preserve TvItemBinding(UBound(TvItemBinding) + 1)                                     '�������ͼ��Ŀ��
                With TvItemBinding(UBound(TvItemBinding))                                                   '��������ͼ��Ŀ��
                    .FileIndex = UBound(CurrentProject.Folders)                                                 '�ļ�������
                    .TVITEM = CreatedTreeItem                                                                   '����ͼ�ڵ�
                    .IsFolder = True                                                                            '���Ϊ�ļ���
                End With
                With CurrentProject.Folders(UBound(CurrentProject.Folders))                                 '������Ŀ��Ϣ����ļ�����Ϣ
                    .FolderPath = FolderPath & NewText                                                          '�ļ���·��
                    If ParentOfCreating = ProjectNameTvItem Then                                                '���ĸ�ڵ�����Ŀ�ڵ� �Ͱ���������Ϊ0��������ĿĿ¼�£�
                        .ParentFolder = 0
                    Else                                                                                        '����ͼ�¼ĸ�ڵ�
                        .ParentFolder = TvItemBinding(i).FileIndex
                    End If
                End With
            Else                                                                                        '����Ǵ����ļ��͸�����Ŀ��Ϣ����ļ���Ϣ
                ReDim Preserve CurrentProject.Files(UBound(CurrentProject.Files) + 1)                       '�����Ŀ��Ϣ����ļ���Ϣ
                ReDim Preserve TvItemBinding(UBound(TvItemBinding) + 1)                                     '�������ͼ��Ŀ��
                With TvItemBinding(UBound(TvItemBinding))                                                   '��������ͼ��Ŀ��
                    .FileIndex = UBound(CurrentProject.Files)                                                   '�ļ�����
                    .TVITEM = CreatedTreeItem                                                                   '����ͼ�ڵ�
                    .IsFolder = False                                                                           '���Ϊ�ļ�
                End With
                With CurrentProject.Files(UBound(CurrentProject.Files))                                     '������Ŀ��Ϣ����ļ���Ϣ
                    .FilePath = ProjectFolderPath & FolderPath & NewText                                        '�ļ�·��
                    If ParentOfCreating = ProjectNameTvItem Then                                                '���ĸ�ڵ�����Ŀ�ڵ� �Ͱ���������Ϊ0��������ĿĿ¼�£�
                        .FolderIndex = 0
                    Else                                                                                        '����ͼ�¼ĸ�ڵ�
                        .FolderIndex = ParentFolderIndex
                    End If
                    .Changed = False                                                                            '����ļ�Ϊδ����
                    .PrevLine = 0
                    ReDim .Breakpoints(0)                                                                       '��ʼ���ļ��ϵ��б�
                End With
            End If
            
            IsCreatingFolder = False
            IsCreatingFile = False
        End If
        Exit Sub
    End If
    
    If hTreeItem = ProjectNameTvItem Then                                                       '����б����Ӧ���ǹ������ƣ�����Ĺ����ļ���
        Err.Clear
        Name ProjectFilePath As ProjectFolderPath & NewText & ".myproj"
        If Err.Number <> 0 Then                                                                     '������ʱ��������
            NoSkinMsgBox Lang_SolutionExplorer_RenameFailure_1 & ProjectFilePath & _
                Lang_SolutionExplorer_RenameFailure_2 & Err.Number & " - " & Err.Description, vbExclamation, Lang_Msgbox_Error
            bCancel = True
        Else                                                                                        '�������ɹ�
            ProjectFilePath = ProjectFolderPath & NewText & ".myproj"                                   '���¹����ļ�·��
            CurrentProject.ProjectName = NewText                                                        '���¹�������
            frmMain.Caption = NewText & " - " & Lang_Application_Title                                  '���������ڱ���
            CurrentProject.Changed = True                                                               '��ǹ����Ѹ���
        End If
        Exit Sub
    End If
    For i = 0 To UBound(TvItemBinding)                                                          '�����б����Ӧ���ļ����
        If hTreeItem = TvItemBinding(i).TVITEM Then                                                 '�ҵ�ƥ����ļ��ͽ���������
            If TvItemBinding(i).IsFolder Then                                                           '���ѡ�����Ŀ���ļ���
                With CurrentProject.Folders(TvItemBinding(i).FileIndex)
                    Err.Clear
                    If .ParentFolder = 0 Then                                                                    '������ڸ�Ŀ¼�£��Ͳ���Ҫ��·���мӡ�\��
                        Name ProjectFolderPath & .FolderPath As ProjectFolderPath & NewText
                    Else
                        Name ProjectFolderPath & .FolderPath As ProjectFolderPath & CurrentProject.Folders(.ParentFolder).FolderPath & "\" & NewText
                    End If
                    If Err.Number <> 0 Then                                                                     '������ʱ��������
                        MsgBox "Error!"     'todo
                        bCancel = True
                    Else                                                                                        '�������ɹ�
                        If .ParentFolder = 0 Then                                                               '�������·��
                            .FolderPath = NewText
                        Else
                            .FolderPath = CurrentProject.Folders(.ParentFolder).FolderPath & "\" & NewText
                        End If
                        RenameFolder TvItemBinding(i).FileIndex                                                     '��������ڵ��������ӽڵ��·��
                    End If
                End With
            Else                                                                                        '���ѡ�����Ŀ���ļ�
                With CurrentProject.Files(TvItemBinding(i).FileIndex)
                    Err.Clear
                    If .FolderIndex = 0 Then                                                                    '������ڸ�Ŀ¼�£��Ͳ���Ҫ��·���мӡ�\��
                        Name (.FilePath) As ProjectFolderPath & NewText
                    Else
                        Name (.FilePath) As ProjectFolderPath & CurrentProject.Folders(.FolderIndex).FolderPath & "\" & NewText
                    End If
                    If Err.Number <> 0 Then                                                                     '������ʱ��������
                        NoSkinMsgBox Lang_SolutionExplorer_RenameFailure_1 & .FilePath & _
                            Lang_SolutionExplorer_RenameFailure_2 & Err.Number & " - " & Err.Description, vbExclamation, Lang_Msgbox_Error
                        bCancel = True
                    Else                                                                                        '�������ɹ�
                        .TargetWindow.Caption = NewText                                                         'ˢ�´��ڱ���
                        frmMain.TabBar.UpdateCaptions
                        If .FolderIndex = 0 Then                                                                '�����ļ�·��
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
    bCancel = True                                                                              '��ʵӦ�ò����Ҳ�����Ӧ���ļ��������������Ҳ�����ȡ��������
End Sub

Public Sub SolutionTreeView_ItemExpanding(ByVal hTreeItem As Long, bCancel As Boolean)

End Sub

Public Sub SolutionTreeView_KeyDown(ByVal KeyCode As Long, ByVal IsLongPress As Boolean)
    If KeyCode = vbKeyF2 Then                                                                   '��ӦF2��: ������
        Me.SolutionTreeView.EditLabel Me.SolutionTreeView.GetSelectedItem()
    ElseIf KeyCode = VK_APPS Then                                                               '��Ӧ�˵���: �����˵�
        Call SolutionTreeView_RightClick(True)
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

'����:      bPopupMenuFromItem: �Ƿ�����б����λ�õ����˵������ڴ���˵���
Public Sub SolutionTreeView_RightClick(bPopupMenuFromItem As Boolean)
    Dim i               As Long
    Dim hTreeItem       As Long
    Dim ItemRect        As RECT
    Dim WindowRect      As RECT
    
    hTreeItem = Me.SolutionTreeView.GetSelectedItem()
    If bPopupMenuFromItem Then                                                                  '����Ǹ��ݸ����б����λ�õ����˵����ͻ�ȡ�б����λ��
        CopyMemory ItemRect, hTreeItem, ByVal 4                                                     '*(HTREEITEM*)&ItemRect = hTreeItem
        SendMessageA Me.SolutionTreeView.TreeViewHwnd, TVM_GETITEMRECT, ByVal 0, ByVal VarPtr(ItemRect)
        GetWindowRect Me.SolutionTreeView.TreeViewHwnd, WindowRect
        ItemRect.Left = WindowRect.Left * Screen.TwipsPerPixelX                                     '������б����������Ļ�ϵ�����
        ItemRect.bottom = (ItemRect.bottom + WindowRect.Top) * Screen.TwipsPerPixelY
    Else                                                                                        '����ʹ�ò˵���Ĭ�ϵ���λ��
        ItemRect.Left = -1
        ItemRect.bottom = -1
    End If
    
    '�ж�ѡ�����б�������Ͳ�������Ӧ�Ĳ˵�
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
        Me.mnuProjectItemPopup.PopupMenu 0, CSng(ItemRect.Left), CSng(ItemRect.bottom)
    Else                                                                                        '������ѡ�����Ŀ�Ƿ�Ϊ�����ļ�
        For i = 0 To UBound(TvItemBinding)
            If hTreeItem = TvItemBinding(i).TVITEM Then                                             '������ҵ���Ӧ���ļ���˵��ѡ����б������ļ���������Ŀ�ڵ�
                If TvItemBinding(i).IsFolder Then                                                       '���ѡ����б������ļ���
                    Me.mnuProjectItemPopup.PopupMenu 0, CSng(ItemRect.Left), CSng(ItemRect.bottom)
                    Exit Sub
                Else                                                                                    '���ѡ����б������ļ�
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

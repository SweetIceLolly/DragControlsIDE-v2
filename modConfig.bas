Attribute VB_Name = "modConfig"
'====================================================
'����:      �ṩ��д���������ļ��������������á����ԡ��û�ϰ�ߵȺ���
'����:      ����
'�ļ�:      modConfig.bas
'====================================================

Option Explicit

'����ϵ���Ϣ�ṹ
Public Type BreakpointInfo
    CodeLn                                      As Long                                     '�ϵ��Ӧ�Ĵ�����
    Enabled                                     As Boolean                                  '�ϵ��Ƿ񼤻�
    ListViewIndex                               As Long                                     '�ڶϵ��б����е��б���ж�Ӧ���б������
End Type

'����gdb�ϵ���Ϣ�ṹ��ÿ��gdb�ϵ㶼Ӧ�ö�Ӧ��CurrentProject.Files(FileIndex).Breakpoints(BreakpointIndex)
Public Type GdbBreakpointMapInfo
    FileIndex                                   As Long                                     '��Ӧ���ļ����
    BreakpointIndex                             As Long                                     '��Ӧ�Ķϵ����
End Type

'�����ļ�����Ϣ�ṹ
Public Type ProjectFolderStruct
    FolderPath                                  As String                                   '��Ŀ·����ProjectFolderPath�����ļ��е����·�������ԡ�\����β��
    ParentFolder                                As Long                                     'ĸ�ļ�����CurrentProject.Folders��������0����û�У�
End Type

'��������ļ���Ϣ�ṹ
Public Type SourceFileStruct
    PrevLine                                    As Long                                     '����ʱ���ڵ��к�
    Changed                                     As Boolean                                  '�ļ��Ƿ񱻸���
    FilePath                                    As String                                   '�ļ�·��
    FolderIndex                                 As Long                                     '����Ŀ¼��CurrentProject.Folders������
    TargetWindow                                As frmCodeWindow                            '��Ӧ�Ĵ��봰�壬ÿ�����е�ʱ�򶼻᲻һ��
    Breakpoints()                               As BreakpointInfo                           '���жϵ���Ϣ
End Type

'���屣��ר�õĴ����ļ���Ϣ�ṹ
Public Type SourceFileStruct_Save
    PrevLine                                    As Long                                     '����ʱ���ڵ��к�
    FileName                                    As String                                   '�ļ����ƣ������·����
    FolderIndex                                 As Long                                     '����Ŀ¼��CurrentProject.Folders������
    Breakpoints()                               As BreakpointInfo                           '���жϵ���Ϣ
End Type

'���幤���ļ��ṹ
Public Type ProjectFileStruct
    ProjectName                                 As String                                   '��������
    ProjectType                                 As Integer                                  '�������͡����frmMain��ProjectType������˵��
    Changed                                     As Boolean                                  '�ļ��Ƿ񱻸���
    Files()                                     As SourceFileStruct                         '���̰����������ļ�
    Folders()                                   As ProjectFolderStruct                      '���̰����������ļ���
End Type

'���屣��ר�õĹ����ļ��ṹ
Public Type ProjectFileStruct_Save
    ProjectName                                 As String                                   '��������
    ProjectType                                 As Integer                                  '�������͡����frmMain��ProjectType������˵��
    Files()                                     As SourceFileStruct_Save                    '���̰����������ļ�
    Folders()                                   As ProjectFolderStruct                      '���̰����������ļ��У���������0��
End Type

'��������ͼ�б������ļ���ţ���ProjectFileStruct.Files���������󶨵Ľṹ
Public Type TvItemToFileIndex
    TVITEM                                      As Long                                     '�ļ���Ŷ�Ӧ������ͼ�б���
    FileIndex                                   As Long                                     '����ͼ�б����Ӧ���ļ���Ż����ļ�������
    IsFolder                                    As Boolean                                  '��Ӧ����Ŀ�Ƿ�Ϊ�ļ���
End Type

'===================================================================
'�����ı������������Ƕ����ԣ���ʹ�ñ���������ÿһ�����ֵ��ַ���
Public Lang_Msgbox_Error                        As String
Public Lang_Msgbox_Confirm                      As String

Public Lang_TitleBar_Max                        As String
Public Lang_TitleBar_Restore                    As String
Public Lang_TitleBar_Min                        As String
Public Lang_TitleBar_Close                      As String

Public Lang_CodeWindow_Caption                  As String
Public Lang_ControlBox_Caption                  As String
Public Lang_Disassembly_Caption                 As String
Public Lang_ErrorList_Caption                   As String
Public Lang_Immediate_Caption                   As String
Public Lang_Memory_Caption                      As String
Public Lang_Modules_Caption                     As String
Public Lang_Output_Caption                      As String
Public Lang_Properties_Caption                  As String
Public Lang_Registers_Caption                   As String
Public Lang_Threads_Caption                     As String
Public Lang_Watch_Caption                       As String

Public Lang_Create_Caption                      As String
Public Lang_Create_CreateLabel                  As String
Public Lang_Create_RecentLabel                  As String
Public Lang_Create_NewWindowProgram             As String
Public Lang_Create_NewConsoleProgram            As String
Public Lang_Create_NewEmptyCpp                  As String
Public Lang_Create_OpenProject                  As String

Public Lang_CreateOptions_Caption               As String
Public Lang_CreateOptions_ProjectNameLabel      As String
Public Lang_CreateOptions_ProjectFolderLabel    As String
Public Lang_CreateOptions_Browse                As String
Public Lang_CreateOptions_Main_NoArgs           As String
Public Lang_CreateOptions_Main_Args             As String
Public Lang_CreateOptions_WinMain               As String
Public Lang_CreateOptions_Include               As String
Public Lang_CreateOptions_OK                    As String
Public Lang_CreateOptions_Cancel                As String
Public Lang_CreateOptions_ProjectType           As String
Public Lang_CreateOptions_TypeOption_1          As String
Public Lang_CreateOptions_TypeOption_2          As String
Public Lang_CreateOptions_TypeOption_3          As String
Public Lang_CreateOptions_BrowseCaption         As String
Public Lang_CreateOptions_ProjectNameRequired   As String
Public Lang_CreateOptions_InvalidProjectName    As String
Public Lang_CreateOptions_InvalidProjectPath    As String
Public Lang_CreateOptions_CreationFailure_1     As String
Public Lang_CreateOptions_CreationFailure_2     As String
Public Lang_CreateOptions_WindowProgram         As String
Public Lang_CreateOptions_ConsoleProgram        As String
Public Lang_CreateOptions_PlainCPP              As String
Public Lang_CreateOptions_SameNameReplace_1     As String
Public Lang_CreateOptions_SameNameReplace_2     As String
Public Lang_CreateOptions_CreateProjectFailed   As String

Public Lang_Application_Title                   As String
Public Lang_Main_SaveBeforeCompile              As String
Public Lang_Main_SaveFailedBeforeCompile        As String
Public Lang_Main_ReplaceExe_1                   As String
Public Lang_Main_ReplaceExe_2                   As String
Public Lang_Main_StartingGcc                    As String
Public Lang_Main_GccStartFailed                 As String
Public Lang_Main_CompileSucceed                 As String
Public Lang_Main_CompileFailed                  As String
Public Lang_Main_Run_Menu_Start                 As String
Public Lang_Main_Run_Menu_Continue              As String
Public Lang_Main_RunFailed                      As String
Public Lang_Main_RunSucceed                     As String
Public Lang_Main_GdbFailed                      As String
Public Lang_Main_GdbSucceed                     As String
Public Lang_Main_GdbAttaching                   As String
Public Lang_Main_GdbAttachFailed_1              As String
Public Lang_Main_GdbAttachFailed_2              As String
Public Lang_Main_GdbLoadingSymbols_1            As String
Public Lang_Main_GdbLoadingSymbols_2            As String
Public Lang_Main_GdbLoadSymbolsFailure_1        As String
Public Lang_Main_GdbLoadSymbolsFailure_2        As String
Public Lang_Main_DebugAborted                   As String
Public Lang_Main_GdbBreakpointError_1           As String
Public Lang_Main_GdbBreakpointError_2           As String
Public Lang_Main_GdbBreakpointError_3           As String
Public Lang_Main_GdbBreakpoint_Invalid          As String
Public Lang_Main_DebugInfo_1                    As String
Public Lang_Main_DebugInfo_2                    As String
Public Lang_Main_RunningInfo_1                  As String
Public Lang_Main_RunningInfo_2                  As String
Public Lang_Main_Debug_OpenSourceFailure        As String
Public Lang_Main_Debug_BreakpointHit            As String
Public Lang_Main_Debug_Returned                 As String

Public Lang_SolutionExplorer_Caption            As String
Public Lang_SolutionExplorer_RenameFailure_1    As String
Public Lang_SolutionExplorer_RenameFailure_2    As String
Public Lang_SolutionExplorer_NewFolderName      As String
Public Lang_SolutionExplorer_InvalidName        As String

Public Lang_SaveBox_Caption                     As String
Public Lang_SaveBox_Yes                         As String
Public Lang_SaveBox_No                          As String
Public Lang_SaveBox_Cancel                      As String
Public Lang_SaveBox_Prompt                      As String
Public Lang_SaveBox_SaveFailure_1               As String
Public Lang_SaveBox_SaveFailure_2               As String

Public Lang_Breakpoints_Caption                 As String
Public Lang_Breakpoints_ListViewHeader_File     As String
Public Lang_Breakpoints_ListViewHeader_Line     As String
Public Lang_Breakpoints_ListViewHeader_Address  As String
Public Lang_Breakpoints_Info_1                  As String
Public Lang_Breakpoints_Info_2                  As String
Public Lang_Breakpoints_Info_3                  As String
Public Lang_Breakpoints_Info_4                  As String

Public Lang_Locals_Caption                      As String
Public Lang_Locals_Retrieving_Caption           As String
Public Lang_Locals_ListViewHeader_Name          As String
Public Lang_Locals_ListViewHeader_Type          As String
Public Lang_Locals_ListViewHeader_Value         As String
Public Lang_Locals_Error                        As String
Public Lang_Locals_Tooltip_Title                As String

Public Lang_CallStack_Caption                   As String
Public Lang_CallStack_Retrieving_Caption        As String
Public Lang_CallStack_Args                      As String
Public Lang_CallStack_Tooltip_Title             As String
Public Lang_CallStack_NoArg                     As String

Public Lang_ErrorList_Errors                    As String
Public Lang_ErrorList_Warnings                  As String
Public Lang_ErrorList_Info                      As String
Public Lang_ErrorList_Description               As String
Public Lang_ErrorList_File                      As String
Public Lang_ErrorList_Line                      As String
Public Lang_ErrorList_Column                    As String
Public Lang_ErrorList_Tooltip_Title             As String
'===================================================================

Public CurrentProject                           As ProjectFileStruct                        '��ǰ���̵���Ϣ
Public ProjectFolderPath                        As String                                   '��ǰ�����ļ��е�λ�ã���"\"��β��
Public ProjectFilePath                          As String                                   '��ǰ��Ŀ�����ļ���λ��

Public GccPath                                  As String                                   'g++·��
Public GdbPath                                  As String                                   'gdb·��

Public TvItemBinding()                          As TvItemToFileIndex                        '��ǰ���̵�TreeView�б�����ļ���ŵİ�
Public ProjectNameTvItem                        As Long                                     'TreeView�б���͹������Ƶİ�

Public CodeWindows                              As New Collection                           '��ǰ�������еĴ��봰��
Public GdbBreakpoints()                         As GdbBreakpointMapInfo                     '��ǰ�����еĶϵ���Ϣ��gdb��

Public IsExiting                                As Boolean                                  '��ǰ�����Ƿ������˳�

'����:      ��ȡָ��·������ļ����������һ����\����������ݣ�
'����:      strPath: ָ��·��
'����ֵ:    �ָ�������ļ���
Public Function GetFileName(strPath As String) As String
    Dim tmp()               As String
    tmp = Split(strPath, "\")
    GetFileName = tmp(UBound(tmp))
End Function

'����:      ���Ƿ��ļ���
'����:      strName: ��Ҫ�������ļ���
'����ֵ:    True: �Ϸ����ļ���; False: �Ƿ����ļ���
Public Function CheckInvalidFileName(strName As String) As Boolean
    Dim InvalidChars        As String                                                       '�Ƿ��ַ�
    Dim i                   As Integer, j               As Integer
    
    If strName = "." Or strName = ".." Then                                                 '��������Ƿ�Ϊ��.�����ߡ�..��
        CheckInvalidFileName = False
        Exit Function
    End If
    
    InvalidChars = """/\:?<>*|"
    For i = 1 To Len(strName)                                                               '���Ƿ��ַ�
        For j = 1 To Len(InvalidChars)
            If Mid(strName, i, 1) = Mid(InvalidChars, j, 1) Then
                CheckInvalidFileName = False
                Exit Function
            End If
        Next j
    Next i
    
    CheckInvalidFileName = True
End Function

'����:      ����һ���µĴ��봰�ڣ���������ӵ�CodeWindows��
'����:      FileIndex: ���봰�ڶ�Ӧ���ļ����
'����ֵ:    �����Ĵ��봰��
Public Function CreateNewCodeWindow(FileIndex As Long) As frmCodeWindow
    Dim NewCodeWindow       As New frmCodeWindow
    
    NewCodeWindow.FileIndex = FileIndex
    CodeWindows.Add NewCodeWindow, CStr(FileIndex)
    Set CurrentProject.Files(FileIndex).TargetWindow = CodeWindows.Item(CStr(FileIndex))    '�ļ��󶨶�Ӧ�Ĵ��봰�ڡ�ǧ��Ҫ�󶨵�NewCodeWindow��
    Set CreateNewCodeWindow = CodeWindows.Item(CStr(FileIndex))                             '���ش����Ĵ��봰�ڡ�ǧ��Ҫ����NewCodeWindow��
End Function

'����:      ��ȡ��Ӧ���Ե��ַ�����Դ���ú�����ͨ��
'.          �ṩ�ĵ�һ����ԴID������������ַ�������Ӧ��ID
'����:      ResID: ��Ӧ��������Ӧ�ĵ�һ����ԴID���籾����������������Ӧ�ĵ�һ����ԴID��1001
'.          LoadMenuTextOnly: ��ѡ��Ĭ��ΪFalse�����ΪTrue�����ֻ���ز˵��ı���
'.                            ��Ϊ���ز˵��ı���ʹ�õ�frmMain��frmMain�ᱻ���أ�
'.                            ������frmMain��Initialize�¼��в��˼��ز˵��ı�������Ӧ����Load�¼��м���
'����ֵ:    �����ȡ�ɹ�������True�����򷵻�False
Public Function LoadLanguage(ResID As Long, Optional LoadMenuTextOnly As Boolean = False) As Boolean
    On Error Resume Next
    LoadLanguage = True
    
    '��ȡ�˵��ַ���
    If LoadMenuTextOnly Then
        Dim id          As Long
        
        '�����ڲ˵�
        For id = 0 To 69
            frmMain.DarkMenu.MenuText(id) = LoadResString(ResID + id)
            If Err.Number <> 0 Then
                LoadLanguage = False
                Exit Function
            End If
        Next id
        
        '������Դ�����������˵�1
        For id = 1 To 4
            frmSolutionExplorer.mnuItemPopup.MenuText(id) = LoadResString(ResID + 99 + id)
            If Err.Number <> 0 Then
                LoadLanguage = False
                Exit Function
            End If
        Next id
        
        For id = 1 To 13
            frmSolutionExplorer.mnuProjectItemPopup.MenuText(id) = LoadResString(ResID + 199 + id)
            If Err.Number <> 0 Then
                LoadLanguage = False
                Exit Function
            End If
        Next id
        
        Exit Function
    End If
    
    '��ȡ���е��ַ���
    Lang_Msgbox_Error = "����"
    Lang_Msgbox_Confirm = "ȷ��"
    
    Lang_TitleBar_Max = "���"
    Lang_TitleBar_Restore = "��ԭ"
    Lang_TitleBar_Min = "��С��"
    Lang_TitleBar_Close = "�ر�"
    
    Lang_CodeWindow_Caption = "���봰��"
    Lang_ControlBox_Caption = "�ؼ���"
    Lang_Disassembly_Caption = "�����"
    Lang_ErrorList_Caption = "�����б�"
    Lang_Immediate_Caption = "��������"
    Lang_Memory_Caption = "�ڴ�"
    Lang_Modules_Caption = "ģ��"
    Lang_Output_Caption = "���"
    Lang_Properties_Caption = "����"
    Lang_Registers_Caption = "�Ĵ���"
    Lang_Threads_Caption = "�߳�"
    Lang_Watch_Caption = "���Ӵ���"
    
    Lang_Create_Caption = "�½���Ŀ"
    Lang_Create_CreateLabel = "����"
    Lang_Create_RecentLabel = "���"
    Lang_Create_NewWindowProgram = "       �½����ڳ���"
    Lang_Create_NewConsoleProgram = "       �½�����̨����"
    Lang_Create_NewEmptyCpp = "       �½��հ�C++����"
    Lang_Create_OpenProject = "       �򿪹���..."
    
    Lang_CreateOptions_Caption = "�½���Ŀ"
    Lang_CreateOptions_ProjectNameLabel = "��Ŀ����:"
    Lang_CreateOptions_ProjectFolderLabel = "��Ŀ�ļ���:"
    Lang_CreateOptions_Browse = "���..."
    Lang_CreateOptions_Main_NoArgs = "����д��main ���޲�����"
    Lang_CreateOptions_Main_Args = "����д��main ���в�����"
    Lang_CreateOptions_WinMain = "����д��WinMain"
    Lang_CreateOptions_Include = "#include <stdio.h>"
    Lang_CreateOptions_OK = "ȷ��"
    Lang_CreateOptions_Cancel = "ȡ��"
    Lang_CreateOptions_ProjectType = "��Ŀ����"
    Lang_CreateOptions_TypeOption_1 = "���ڳ���"
    Lang_CreateOptions_TypeOption_2 = "����̨����"
    Lang_CreateOptions_TypeOption_3 = "�հ�C++����"
    Lang_CreateOptions_BrowseCaption = "ѡ����Ŀ�ļ���"
    Lang_CreateOptions_ProjectNameRequired = "��������Ŀ���ƣ�"
    Lang_CreateOptions_InvalidProjectName = "��Ч����Ŀ����: "
    Lang_CreateOptions_InvalidProjectPath = "ָ������Ŀ�ļ���·����Ч��"
    Lang_CreateOptions_CreationFailure_1 = "�޷�����"
    Lang_CreateOptions_CreationFailure_2 = " ����ȷ����Ŀ��������Ч�ġ�"
    Lang_CreateOptions_WindowProgram = "�´��ڳ���"
    Lang_CreateOptions_ConsoleProgram = "�¿���̨����"
    Lang_CreateOptions_PlainCPP = "�¿հ�C++����"
    Lang_CreateOptions_SameNameReplace_1 = "ѡ���Ŀ¼�����������ļ�: "
    Lang_CreateOptions_SameNameReplace_2 = "���Ƿ񸲸ǣ�"
    Lang_CreateOptions_CreateProjectFailed = "����������ʧ��"
    
    Lang_Application_Title = "�Ͽؼ���"
    Lang_Main_SaveBeforeCompile = "�Ƿ��ȱ��������ļ��ٽ��б��룿"
    Lang_Main_SaveFailedBeforeCompile = "�����ļ�ʱ���������Ƿ�������б��룿"
    Lang_Main_ReplaceExe_1 = "��⵽�ڱ���Ŀ¼�����ļ��뼴������Ŀ�ִ���ļ�����: "
    Lang_Main_ReplaceExe_2 = " �Ƿ�������룿���ļ����ᱻ���ǡ�"
    Lang_Main_StartingGcc = "��������g++���б���: "
    Lang_Main_GccStartFailed = "�޷�����g++: "
    Lang_Main_CompileSucceed = "�������: EXE·��: "
    Lang_Main_CompileFailed = "����ʧ�ܣ�"
    Lang_Main_Run_Menu_Start = LoadResString(ResID + 52)
    Lang_Main_Run_Menu_Continue = "�������� (&R)"
    Lang_Main_RunFailed = "�޷����� "
    Lang_Main_RunSucceed = "���������Խ���: ����ID: "
    Lang_Main_GdbFailed = "����gdb���Թܵ�ʧ�ܣ��޷����е��ԡ�"
    Lang_Main_GdbSucceed = "����gdb���Թܵ�: ����ID: "
    Lang_Main_GdbAttaching = "���ڸ��ӽ���..."
    Lang_Main_GdbAttachFailed_1 = "gdb���ӵ�����"
    Lang_Main_GdbAttachFailed_2 = "ʧ�ܣ��޷����е��ԡ�"
    Lang_Main_GdbLoadingSymbols_1 = "���ڴ��ļ�"
    Lang_Main_GdbLoadingSymbols_2 = "��ȡ����..."
    Lang_Main_GdbLoadSymbolsFailure_1 = "�ӿ�ִ���ļ�"
    Lang_Main_GdbLoadSymbolsFailure_2 = " ���ط���ʧ�ܣ�����ζ�Ŷϵ㡢�鿴���ر����ȵ��Թ��ܽ��޷������������Ƿ�������ԣ�"
    Lang_Main_DebugAborted = "�������ԡ�"
    Lang_Main_GdbBreakpointError_1 = "�ϵ����: ���ļ� "
    Lang_Main_GdbBreakpointError_2 = " ���Ҳ�����"
    Lang_Main_GdbBreakpointError_3 = "�С�"
    Lang_Main_GdbBreakpoint_Invalid = "<�ϵ���Ч>"
    Lang_Main_RunningInfo_1 = "����"
    Lang_Main_RunningInfo_2 = "��������"
    Lang_Main_Debug_OpenSourceFailure = "�޷��򿪴����ļ�: "
    Lang_Main_Debug_BreakpointHit = "�ϵ�������"
    Lang_Main_Debug_Returned = "�����˳�������: "
    
    Lang_SolutionExplorer_Caption = "������Դ������"
    Lang_SolutionExplorer_RenameFailure_1 = "Ϊ�ļ�"
    Lang_SolutionExplorer_RenameFailure_2 = " ������ʧ��: "
    Lang_SolutionExplorer_NewFolderName = "���ļ���"
    Lang_SolutionExplorer_InvalidName = "��Ч�����ƣ�"
    
    Lang_SaveBox_Caption = "����"
    Lang_SaveBox_Yes = "��"
    Lang_SaveBox_No = "��"
    Lang_SaveBox_Cancel = "ȡ��"
    Lang_SaveBox_Prompt = "�Ƿ񱣴�������ѡ����ļ���"
    Lang_SaveBox_SaveFailure_1 = "�޷������ļ���"
    Lang_SaveBox_SaveFailure_2 = " ���Ƿ�������������ļ���"
    
    Lang_Breakpoints_Caption = "�ϵ��б�"
    Lang_Breakpoints_ListViewHeader_File = "�ļ�"
    Lang_Breakpoints_ListViewHeader_Line = "�к�"
    Lang_Breakpoints_ListViewHeader_Address = "��ַ"
    Lang_Breakpoints_Info_1 = "�ϵ��ڵ�"
    Lang_Breakpoints_Info_2 = "��: "
    Lang_Breakpoints_Info_3 = "������"
    Lang_Breakpoints_Info_4 = "�ѽ���"
    
    Lang_Locals_Caption = "����"
    Lang_Locals_Retrieving_Caption = "���� - ���ڻ�ȡ..."
    Lang_Locals_ListViewHeader_Name = "����"
    Lang_Locals_ListViewHeader_Type = "����"
    Lang_Locals_ListViewHeader_Value = "ֵ"
    Lang_Locals_Error = "<����>"
    Lang_Locals_Tooltip_Title = "���ر�����Ϣ: "
    
    Lang_CallStack_Caption = "���ö�ջ"
    Lang_CallStack_Retrieving_Caption = "���ö�ջ - ���ڻ�ȡ..."
    Lang_CallStack_Args = "����"
    Lang_CallStack_Tooltip_Title = "���ö�ջ��Ϣ:"
    Lang_CallStack_NoArg = "<��>"
    
    Lang_ErrorList_Errors = " ����"
    Lang_ErrorList_Warnings = " ����"
    Lang_ErrorList_Info = " ��Ϣ"
    Lang_ErrorList_Description = "����"
    Lang_ErrorList_File = "�ļ�"
    Lang_ErrorList_Line = "��"
    Lang_ErrorList_Column = "��"
    Lang_ErrorList_Tooltip_Title = "������Ϣ"
End Function


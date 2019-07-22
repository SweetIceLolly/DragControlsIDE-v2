Attribute VB_Name = "modConfig"
'====================================================
'����:      �ṩ��д���������ļ��������������á����ԡ��û�ϰ�ߵȺ���
'����:      ����
'�ļ�:      modConfig.bas
'====================================================

Option Explicit

'����ϵ���Ϣ�ṹ
Public Type BreakpointInfo
    CodeLn                  As Long                                             '�ϵ��Ӧ�Ĵ�����
    Enabled                 As Boolean                                          '�ϵ��Ƿ񼤻�
End Type

'��������ļ���Ϣ�ṹ
Public Type SourceFileStruct
    IsHeaderFile            As Boolean                                          '�Ƿ�Ϊͷ�ļ�
    PrevLine                As Long                                             '����ʱ���ڵ��к�
    Changed                 As Boolean                                          '�ļ��Ƿ񱻸���
    FilePath                As String                                           '�ļ�·��
    TargetWindow            As frmCodeWindow                                    '��Ӧ�Ĵ��봰�壬ÿ�����е�ʱ�򶼻᲻һ��
    Breakpoints()           As BreakpointInfo                                   '���жϵ���Ϣ
End Type

'���屣��ר�õĴ����ļ���Ϣ�ṹ
Public Type SourceFileStruct_Save
    IsHeaderFile            As Boolean                                          '�Ƿ�Ϊͷ�ļ�
    PrevLine                As Long                                             '����ʱ���ڵ��к�
    FileName                As String                                           '�ļ����ƣ������·����
    Breakpoints()           As BreakpointInfo                                   '���жϵ���Ϣ
End Type

'���幤���ļ��ṹ
Public Type ProjectFileStruct
    ProjectName             As String                                           '��������
    ProjectType             As Integer                                          '�������͡����frmMain��ProjectType������˵��
    Changed                 As Boolean                                          '�ļ��Ƿ񱻸���
    Files()                 As SourceFileStruct                                 '���̰����������ļ�
End Type

'���屣��ר�õĹ����ļ��ṹ
Public Type ProjectFileStruct_Save
    ProjectName             As String                                           '��������
    ProjectType             As Integer                                          '�������͡����frmMain��ProjectType������˵��
    Files()                 As SourceFileStruct_Save                            '���̰����������ļ�
End Type

'��������ͼ�б������ļ���Ű󶨵Ľṹ
Public Type TvItemToFileIndex
    TVITEM                  As Long                                             '�ļ���Ŷ�Ӧ������ͼ�б���
    FileIndex               As Long                                             '����ͼ�б����Ӧ���ļ����
End Type

'===================================================================
'�����ı������������Ƕ����ԣ���ʹ�ñ���������ÿһ�����ֵ��ַ���
Public Lang_Msgbox_Error                        As String
Public Lang_Msgbox_Confirm                      As String

Public Lang_TitleBar_Max                        As String
Public Lang_TitleBar_Restore                    As String
Public Lang_TitleBar_Min                        As String
Public Lang_TitleBar_Close                      As String

Public Lang_Breakpoints_Caption                 As String
Public Lang_CallStack_Caption                   As String
Public Lang_CodeWindow_Caption                  As String
Public Lang_ControlBox_Caption                  As String
Public Lang_Disassembly_Caption                 As String
Public Lang_ErrorList_Caption                   As String
Public Lang_Immediate_Caption                   As String
Public Lang_Locals_Caption                      As String
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
Public Lang_CreateOptions_BrowseCaption         As String
Public Lang_CreateOptions_InvalidProjectPath    As String
Public Lang_CreateOptions_NameConflict_1        As String
Public Lang_CreateOptions_NameConflict_2        As String
Public Lang_CreateOptions_CreationFailure_1     As String
Public Lang_CreateOptions_CreationFailure_2     As String
Public Lang_CreateOptions_SourceFile            As String
Public Lang_CreateOptions_WindowProgram         As String
Public Lang_CreateOptions_ConsoleProgram        As String
Public Lang_CreateOptions_PlainCPP              As String

Public Lang_Application_Title                   As String
Public Lang_Main_SaveBeforeCompile              As String
Public Lang_Main_SaveFailedBeforeCompile        As String
Public Lang_Main_ReplaceExe_1                   As String
Public Lang_Main_ReplaceExe_2                   As String
Public Lang_Main_StartingGcc                    As String
Public Lang_Main_GccStartFailed                 As String
Public Lang_Main_CompileSucceed                 As String
Public Lang_Main_CompileFailed                  As String
Public Lang_Main_RunFailed                      As String
Public Lang_Main_GdbFailed                      As String
Public Lang_Main_GdbAttachFailed_1              As String
Public Lang_Main_GdbAttachFailed_2              As String
Public Lang_Main_GdbLoadSymbolsFailure_1        As String
Public Lang_Main_GdbLoadSymbolsFailure_2        As String
Public Lang_Main_DebugAborted                   As String
Public Lang_Main_DebugInfo_1                    As String
Public Lang_Main_DebugInfo_2                    As String

Public Lang_SolutionExplorer_Caption            As String
Public Lang_SolutionExplorer_RenameFailure_1    As String
Public Lang_SolutionExplorer_RenameFailure_2    As String

Public Lang_SaveBox_Caption                     As String
Public Lang_SaveBox_Yes                         As String
Public Lang_SaveBox_No                          As String
Public Lang_SaveBox_Cancel                      As String
Public Lang_SaveBox_Prompt                      As String
Public Lang_SaveBox_SaveFailure_1               As String
Public Lang_SaveBox_SaveFailure_2               As String
'===================================================================

Public CurrentProject       As ProjectFileStruct                                '��ǰ���̵���Ϣ
Public ProjectFolderPath    As String                                           '��ǰ�����ļ��е�λ�ã���"\"��β��
Public ProjectFilePath      As String                                           '��ǰ��Ŀ�����ļ���λ��
Public TvItemBinding()      As TvItemToFileIndex                                '��ǰ���̵�TreeView�б�����ļ���ŵİ�
Public ProjectNameTvItem    As Long                                             'TreeView�б���͹������Ƶİ�
Public CodeWindows          As New Collection                                   '��ǰ�������еĴ��봰��
Public IsExiting            As Boolean                                          '��ǰ�����Ƿ������˳�

'����:      ��ȡָ��·������ļ����������һ����\����������ݣ�
'����:      strPath: ָ��·��
'����ֵ:    �ָ�������ļ���
Public Function GetFileName(strPath As String) As String
    Dim tmp()               As String
    tmp = Split(strPath, "\")
    GetFileName = tmp(UBound(tmp))
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
        
        For id = 0 To 69
            frmMain.DarkMenu.MenuText(id) = LoadResString(ResID + id)
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
    
    Lang_Breakpoints_Caption = "�ϵ��б�"
    Lang_CallStack_Caption = "���ö�ջ"
    Lang_CodeWindow_Caption = "���봰��"
    Lang_ControlBox_Caption = "�ؼ���"
    Lang_Disassembly_Caption = "�����"
    Lang_ErrorList_Caption = "�����б�"
    Lang_Immediate_Caption = "��������"
    Lang_Locals_Caption = "����"
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
    Lang_CreateOptions_BrowseCaption = "ѡ����Ŀ�ļ���"
    Lang_CreateOptions_InvalidProjectPath = "ָ������Ŀ�ļ���·����Ч��"
    Lang_CreateOptions_NameConflict_1 = "��⵽ͬ���ļ�: "
    Lang_CreateOptions_NameConflict_2 = " ���Ƿ����������Ŀ��Ŀ���ļ����ᱻ���ǣ�"
    Lang_CreateOptions_CreationFailure_1 = "�޷�����"
    Lang_CreateOptions_CreationFailure_2 = " ����ȷ����Ŀ��������Ч�ġ�"
    Lang_CreateOptions_SourceFile = "Դ�ļ�"
    Lang_CreateOptions_WindowProgram = "�´��ڳ���"
    Lang_CreateOptions_ConsoleProgram = "�¿���̨����"
    Lang_CreateOptions_PlainCPP = "�¿հ�C++����"
    
    Lang_Application_Title = "�Ͽؼ���"
    Lang_Main_SaveBeforeCompile = "�Ƿ��ȱ��������ļ��ٽ��б��룿"
    Lang_Main_SaveFailedBeforeCompile = "�����ļ�ʱ���������Ƿ�������б��룿"
    Lang_Main_ReplaceExe_1 = "��⵽�ڱ���Ŀ¼�����ļ��뼴������Ŀ�ִ���ļ�����: "
    Lang_Main_ReplaceExe_2 = " �Ƿ�������룿���ļ����ᱻ���ǡ�"
    Lang_Main_StartingGcc = "��������g++���б���..."
    Lang_Main_GccStartFailed = "�޷�����g++��"
    Lang_Main_CompileSucceed = "�������: EXE·��: "
    Lang_Main_CompileFailed = "����ʧ�ܣ�"
    Lang_Main_RunFailed = "�޷����� "
    Lang_Main_GdbFailed = "����gdb���Թܵ�ʧ�ܣ��޷����е��ԡ�"
    Lang_Main_GdbAttachFailed_1 = "gdb���ӵ�����"
    Lang_Main_GdbAttachFailed_2 = "ʧ�ܣ��޷����е��ԡ�"
    Lang_Main_GdbLoadSymbolsFailure_1 = "�ӿ�ִ���ļ�"
    Lang_Main_GdbLoadSymbolsFailure_2 = " ���ط���ʧ�ܣ�����ζ�Ŷϵ㡢�鿴���ر����ȵ��Թ��ܽ��޷������������Ƿ�������ԣ�"
    Lang_Main_DebugAborted = "�������ԡ�"
    Lang_Main_DebugInfo_1 = "�������ڽ���: gdb.exe ����ID: "
    Lang_Main_DebugInfo_2 = " ����ID: "
    
    Lang_SolutionExplorer_Caption = "������Դ������"
    Lang_SolutionExplorer_RenameFailure_1 = "Ϊ�ļ�"
    Lang_SolutionExplorer_RenameFailure_2 = " ������ʧ��: "
    
    Lang_SaveBox_Caption = "����"
    Lang_SaveBox_Yes = "��"
    Lang_SaveBox_No = "��"
    Lang_SaveBox_Cancel = "ȡ��"
    Lang_SaveBox_Prompt = "�Ƿ񱣴�������ѡ����ļ���"
    Lang_SaveBox_SaveFailure_1 = "�޷������ļ���"
    Lang_SaveBox_SaveFailure_2 = " ���Ƿ�������������ļ���"
End Function

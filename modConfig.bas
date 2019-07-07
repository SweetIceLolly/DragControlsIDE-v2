Attribute VB_Name = "modConfig"
'====================================================
'����:      �ṩ��д���������ļ��������������á����ԡ��û�ϰ�ߵȺ���
'����:      ����
'�ļ�:      modConfig.bas
'====================================================

Option Explicit

'����cpp�ļ���Ϣ�ṹ
Public Type SourceFile
    IsHeaderFile            As Boolean                                          '�Ƿ�Ϊͷ�ļ�
    PrevLine                As Long                                             '����ʱ���ڵ��к�
    Changed                 As Boolean                                          '�ļ��Ƿ񱻸���
    FilePath                As String                                           '�ļ�·��
    TargetWindow            As frmCodeWindow                                    '��Ӧ�Ĵ��봰�壬ÿ�����е�ʱ�򶼻᲻һ��
End Type

'���幤���ļ��ṹ
Public Type ProjectFileStruct
    ProjectName             As String                                           '��������
    ProjectType             As Integer                                          '�������͡����frmMain��ProjectType������˵��
    Changed                 As Boolean                                          '�ļ��Ƿ񱻸���
    Files()                 As SourceFile                                       '���̰����������ļ�
End Type

'��������ͼ�б������ļ���Ű󶨵Ľṹ
Public Type TvItemToFileIndex
    TVITEM                  As Long                                             '�ļ���Ŷ�Ӧ������ͼ�б���
    FileIndex               As Long                                             '����ͼ�б����Ӧ���ļ����
End Type

Public CurrentProject       As ProjectFileStruct                                '��ǰ���̵���Ϣ
Public ProjectFolderPath    As String                                           '��ǰ�����ļ��е�λ�ã���"\"��β��
Public ProjectFilePath      As String                                           '��ǰ��Ŀ�����ļ���λ��
Public TvItemBinding()      As TvItemToFileIndex                                '��ǰ���̵�TreeView�б�����ļ���ŵİ�
Public CodeWindows          As New Collection                                   '��ǰ�������еĴ��봰��
Public IsExiting            As Boolean                                          '��ǰ�����Ƿ������˳�

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
'����ֵ:    �����ȡ�ɹ�������True�����򷵻�False
Public Function LoadLanguage(ResID As Long) As Boolean
    On Error Resume Next
    LoadLanguage = True
    
    '��ȡ�˵��ַ���
    Dim id          As Long
    
    For id = 0 To 69
        frmMain.DarkMenu.MenuText(id) = LoadResString(ResID + id)
        If Err.Number <> 0 Then
            LoadLanguage = False
            Exit Function
        End If
    Next id
End Function

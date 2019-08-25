Attribute VB_Name = "modMain"
'====================================================
'����:      �ṩһЩȫ��ͨ�õĺ������細����Ϣ�����
'����:      ����
'�ļ�:      modMain.bas
'====================================================

Option Explicit

'����ϵͳ����Ϣ�������
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'��ȡϵͳ������Ϣ
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

'���ַ���ת���ֽ�����
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
'���ֽ�����ת���ַ���
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Public DebugProgramInfo     As PROCESS_INFORMATION                                      '���ڵ����еĽ�����Ϣ

Public bpRedrawFileIndex    As Long                                                     '��Ҫ�ػ�ϵ�Ĵ��봰������Ӧ���ļ����
    
'����:      ��ȡ�������·�������·������ȱ��"\"�����Զ�����
'����ֵ:    ��"\"��β��·��
Public Function GetAppPath() As String
    GetAppPath = App.Path
    If Right(GetAppPath, 1) <> "\" Then
        GetAppPath = GetAppPath & "\"
    End If
End Function

'����:      ���ַ���ת�����ֽ�����
'����:      strInput: ��Ҫת�����ַ���
'.          AutoAddNullChar: ��ѡ�ġ��Ƿ��Զ����ַ���ĩβ���'\0'��Ĭ��ΪTrue
'����ֵ:    ת���������ֽ�����
Public Function StrConvEx(ByVal strInput As String, Optional AutoAddNullChar As Boolean = True) As Byte()
    Dim nBytes      As Long
    Dim tmpBuf()    As Byte
    
    If AutoAddNullChar Then
        strInput = strInput & vbNullChar                                                        '���ַ���ĩβ����'\0'
    End If
    nBytes = WideCharToMultiByte(CP_ACP, 0, ByVal StrPtr(strInput), -1, 0, 0, 0, 0)         '��ȡ��Ҫ�Ļ�������С
    ReDim tmpBuf(nBytes - 1)                                                                '���仺����
    WideCharToMultiByte CP_ACP, 0, ByVal StrPtr(strInput), -1, _
        ByVal VarPtr(tmpBuf(0)), nBytes - 1, 0, 0                                           'ת��
    If Not AutoAddNullChar Then                                                             '����û�ָ�����Զ����'\0'��ȥ��ĩβ��'\0'
        ReDim Preserve tmpBuf(UBound(tmpBuf) - 1)
    End If
    StrConvEx = tmpBuf
End Function

'����:      ���ֽ�����ת�����ַ���
'����:      ByteArrInput: ��Ҫת�����ֽ�����
'����ֵ:    ת���������ַ���
Public Function ByteArrayConv(ByteArrInput() As Byte) As String
    Dim nBytes      As Long                                                                                     '��������Ҫ����Ĵ�С
    Dim tmpStr      As String                                                                                   '�����ַ���
    Dim NullCharPos As Long                                                                                     ''\0'���ַ����е�λ��
    
    nBytes = MultiByteToWideChar(CP_ACP, 0, ByVal VarPtr(ByteArrInput(0)), UBound(ByteArrInput) + 1, 0, 0)      '��ȡ��Ҫ�Ļ�������С
    tmpStr = String(nBytes, vbNullChar)                                                                         '���仺����
    nBytes = MultiByteToWideChar(CP_ACP, 0, ByVal VarPtr(ByteArrInput(0)), _
        UBound(ByteArrInput) + 1, ByVal StrPtr(tmpStr), nBytes)                                                 'ת��
    NullCharPos = InStr(tmpStr, vbNullChar)
    If NullCharPos > 0 Then
        ByteArrayConv = Left(tmpStr, NullCharPos - 1)
    Else
        ByteArrayConv = tmpStr
    End If
End Function

'����:      �жϽ������Ƿ������ָ��PID�Ľ���
'����:      hProcess: ���̾��
'����ֵ:    ָ���Ľ����Ƿ����
Public Function ProcessExists(ByVal hProcess As Long) As Boolean
    Dim ret         As Long
    
    ret = WaitForSingleObject(hProcess, 0)                                                  '�жϽ����Ƿ��˳�
    ProcessExists = (ret = WAIT_TIMEOUT)                                                    '������ֵΪ��ʱ˵��������������
End Function

'����:      ��ʹ��MsgBoxǰ�ȹر�Ƥ�����������MsgBox���ѿ� :)
'����:      MsgBox��ǰ��������
'����ֵ:    MsgBox�ķ���ֵ
Public Function NoSkinMsgBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String) As VbMsgBoxResult
    frmMain.SkinFramework.AutoApplyNewThreads = False
    frmMain.SkinFramework.AutoApplyNewWindows = False
    NoSkinMsgBox = MsgBox(Prompt, Buttons, Title)
    frmMain.SkinFramework.AutoApplyNewThreads = True
    frmMain.SkinFramework.AutoApplyNewWindows = True
End Function

'����:      ������16λ�������ϳ�һ��32λ��Long����
'����:      wLow, wHigh: �ֱ��ǵ�16λ�͸�16λ
'����ֵ:    �ϳɵ���
Public Function MakeLong(wLow As Long, wHigh As Long) As Long
    MakeLong = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

'����:      ��ȡһ��32λ���ĸ�16λ
'����:      lValue: ��ֵ
'����ֵ:    ��16λ����ֵ
Public Function HiWord(lValue As Long) As Integer
    If lValue And &H80000000 Then
        HiWord = (lValue \ 65535) - 1
    Else
        HiWord = lValue \ 65535
    End If
End Function

'����:      ��ȡһ��32λ���ĵ�16λ
'����:      lValue: ��ֵ
'����ֵ:    ��16λ����ֵ
Public Function LoWord(lValue As Long) As Integer
    If lValue And &H8000& Then
        LoWord = &H8000 Or (lValue And &H7FFF&)
    Else
        LoWord = lValue And &HFFFF&
    End If
End Function

'����:      ͨ��wParam�����Shiftֵ
'����:      wParam: wParamֵ
'����ֵ:    Shiftֵ
Public Function GetShiftValue(wParam As Long) As Long
    GetShiftValue = (wParam And MK_CONTROL) Or (wParam And MK_SHIFT)
End Function

'����:      �޸����������ȫ���������������Ҽ��˵��޷��رյ�����
'����:      hWnd: ���ھ��
'.          uMsg: ��Ϣֵ
'.          wParam, lParam: ��Ϣ�Ĳ���
'����ֵ:    ��Ϣ������ֵ
Public Function MainWindowMaximizeCloseFixProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_GETMINMAXINFO Then                                                         '���ڳ��Ի�ȡ�����С����Ϣ
        Dim mmi             As MINMAXINFO                                                       '�����С����Ϣ
        Dim rectWorkArea    As RECT                                                             '��Ļ��������С
        
        'lParamΪָ��MINMAXINFO��ָ��
        CopyMemory mmi, ByVal lParam, ByVal Len(mmi)
        SystemParametersInfo SPI_GETWORKAREA, ByVal 0, rectWorkArea, ByVal 0                    '��ȡ��Ļ��������С
        mmi.ptMaxSize.Y = rectWorkArea.bottom - rectWorkArea.Top
        CopyMemory ByVal lParam, mmi, ByVal Len(mmi)                                            '���������Ϣ�еĴ�С��Ϣ���޸�������󻯵�ʱ���ȫ��������
        
        MainWindowMaximizeCloseFixProc = 0                                                      '���������Ϣ֮����Ҫ����0
        Exit Function
    ElseIf uMsg = WM_SYSCOMMAND Then                                                        '��������ʹ���Ҽ��˵��ر�
        If wParam = SC_CLOSE Then
            Dim WindowObj   As Object                                                           '��Ӧ�Ĵ������
            
            CopyMemory ByVal VarPtr(WindowObj), GetPropA(hWnd, "WindowObj"), ByVal 4            '��ȡ�ô��ڶ�Ӧ��Form
            Unload WindowObj                                                                    'ж��Form
        End If
    End If
    MainWindowMaximizeCloseFixProc = CallWindowProc(GetPropA(hWnd, "PrevWndProc"), hWnd, uMsg, wParam, lParam)
End Function

'����:      �ڴ����ı����ػ��ͬʱ�ػ�ϵ�
'����:      hWnd: ���ھ��
'.          uMsg: ��Ϣֵ
'.          wParam, lParam: ��Ϣ�Ĳ���
'����ֵ:    ��Ϣ������ֵ
Public Function EditBreakpointsRedrawProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_PAINT Then                                                                 '���ص�WM_PAINT��Ϣ��ʱ��˳���ػ�ϵ�
        bpRedrawFileIndex = GetPropA(hWnd, "FileIndex")
    End If
    EditBreakpointsRedrawProc = CallWindowProc(GetPropA(hWnd, "PrevWndProc"), hWnd, uMsg, wParam, lParam)
End Function

'����:      �ڡ����ء����ڵ�ListView���б�ͷ������С��ʱ�����ͼƬ��Ŀ��
'����:      hWnd: ���ھ��
'.          uMsg: ��Ϣֵ
'.          wParam, lParam: ��Ϣ�Ĳ���
'����ֵ:    ��Ϣ������ֵ
Public Function LocalsColumnHeaderLayoutProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    
    If uMsg = HDM_LAYOUT Then                                                               '���ص�HDM_LAYOUT��Ϣ��ʱ�����ͼƬ����
        Dim ItemRect        As RECT                                                             '��һ���б�ͷ�Ŀ��
        
        SendMessageA hWnd, HDM_GETITEMRECT, ByVal 0, ByVal VarPtr(ItemRect)                     '��ȡ��һ���б�ͷ�Ŀ��
        ItemRect.Left = (ItemRect.Right - ItemRect.Left) * Screen.TwipsPerPixelX                '�������ȣ�羣�����ֱ�Ӵ����ItemRect.Left
        
        '���㹻�Ŀ�ȾͰ�ͼƬ��Ŀ������Ϊ300��û���㹻�Ŀ�Ⱦ���ͼƬ��Ŀ�������б�ͷ�Ŀ�ȱ仯
        frmLocals.picSelMargin.Width = IIf(ItemRect.Left > frmLocals.picSelMargin.Width, 300, ItemRect.Left)
    End If
    LocalsColumnHeaderLayoutProc = CallWindowProc(GetPropA(hWnd, "PrevWndProc"), hWnd, uMsg, wParam, lParam)
End Function

'����:      �������ء����ڵ�ListView�ػ��ʱ���ػ�ڵ�ͼ��
'����:      hWnd: ���ھ��
'.          uMsg: ��Ϣֵ
'.          wParam, lParam: ��Ϣ�Ĳ���
'����ֵ:    ��Ϣ������ֵ
Public Function LocalsListViewNodesRedrawProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_PAINT Then
        Call frmLocals.RedrawNodeIcons
    End If
    LocalsListViewNodesRedrawProc = CallWindowProc(GetPropA(hWnd, "PrevWndProc"), hWnd, uMsg, wParam, lParam)
End Function

'����:      ��ʾ���򿪡�ͨ�öԻ���
'����:      hWnd: ���øú����Ĵ��ھ��
'.          Title: �Ի������
'.          Filter: �ļ���������ʹ��vbNullChar������ÿ��������
'����ֵ:    �������ȡ�����߳�������""�����򷵻�ѡ����ļ�·��
Public Function ShowOpen(hWnd As Long, Title As String, Filter As String) As String
    Dim ofn                 As OPENFILENAME                                                 '�Ի�����Ϣ
    
    Filter = Filter & vbNullChar
    With ofn                                                                                '���öԻ�����Ϣ
        .lStructSize = Len(ofn)
        .hWndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrFilter = Filter
        .lpstrFile = String(MAX_PATH, vbNullChar)                                               '�����ļ���������
        .nMaxFile = MAX_PATH
        .lpstrFileTitle = Space(MAX_PATH - 1)
        .nMaxFileTitle = MAX_PATH
        .lpstrTitle = Title
        .Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
        .nFilterIndex = 0
    End With
    
    frmMain.SkinFramework.AutoApplyNewThreads = False                                       '��ʱ����Ƥ���ؼ�
    frmMain.SkinFramework.AutoApplyNewWindows = False
    If GetOpenFileNameW(ofn) = 1 Then                                                       '��ʾ����Ի���
        ShowOpen = Split(ofn.lpstrFile, vbNullChar)(0)                                          '��'\0'���ض��ַ���
    End If
    frmMain.SkinFramework.AutoApplyNewThreads = True                                        '��������Ƥ���ؼ�
    frmMain.SkinFramework.AutoApplyNewWindows = True
End Function

'����:      ��ʾ�����桱ͨ�öԻ���
'����:      hWnd: ���øú����Ĵ��ھ��
'.          DefaultName: Ĭ�ϵ��ļ���
'.          Title: �Ի������
'.          Filter: �ļ���������ʹ��vbNullChar������ÿ��������
'����ֵ:    �������ȡ�����߳�������""�����򷵻�ѡ����ļ�·��
Public Function ShowSave(hWnd As Long, DefaultName As String, Title As String, Filter As String) As String
    Dim ofn                 As OPENFILENAME                                                 '�Ի�����Ϣ
    
    DefaultName = DefaultName & String(MAX_PATH - Len(DefaultName), vbNullChar)             '�ַ�����β�����㹻������'\0'����Ϊ������
    Filter = Filter & vbNullChar                                                            '�ַ���ĩβ������'\0'
    With ofn                                                                                '���öԻ�����Ϣ
        .lStructSize = Len(ofn)
        .hWndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrFilter = Filter
        .lpstrFile = DefaultName
        .nMaxFile = MAX_PATH
        .lpstrFileTitle = Space(MAX_PATH - 1)
        .nMaxFileTitle = MAX_PATH
        .lpstrTitle = Title
        .Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
        .nFilterIndex = 0
    End With
    
    frmMain.SkinFramework.AutoApplyNewThreads = False                                       '��ʱ����Ƥ���ؼ�
    frmMain.SkinFramework.AutoApplyNewWindows = False
    If GetSaveFileNameW(ofn) = 1 Then                                                       '��ʾ����Ի���
        ShowSave = Split(ofn.lpstrFile, vbNullChar)(0)                                          '��'\0'���ض��ַ���
    End If
    frmMain.SkinFramework.AutoApplyNewThreads = True                                        '��������Ƥ���ؼ�
    frmMain.SkinFramework.AutoApplyNewWindows = True
End Function

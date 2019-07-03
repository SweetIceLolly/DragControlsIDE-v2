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

Public DebugProgramInfo     As PROCESS_INFORMATION                                      '���ڵ����еĽ�����Ϣ
    
'����:      ��ȡ�������·�������·������ȱ��"\"�����Զ�����
'����ֵ:    ��"\"��β��·��
Public Function GetAppPath() As String
    GetAppPath = App.Path
    If Right(GetAppPath, 1) <> "\" Then
        GetAppPath = GetAppPath & "\"
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

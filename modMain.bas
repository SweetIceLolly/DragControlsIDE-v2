Attribute VB_Name = "modMain"
'====================================================
'����:      �ṩһЩȫ��ͨ�õĺ������細����Ϣ�����
'����:      ����
'�ļ�:      modMain.bas
'====================================================

Option Explicit

'����ϵͳ����Ϣ�������
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
    ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'��ȡϵͳ������Ϣ
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    
'����:      ��ȡ�������·�������·������ȱ��"\"�����Զ�����
'����ֵ:    ��"\"��β��·��
Public Function GetAppPath() As String
    GetAppPath = App.Path
    If Right(GetAppPath, 1) <> "\" Then
        GetAppPath = GetAppPath & "\"
    End If
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

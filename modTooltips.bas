Attribute VB_Name = "modTooltips"
'====================================================
'����:      ������ʾ�ı�ģ�飬����֧�ֶ��еġ������Զ�����ʽ�Ĺ�����ʾ�ı�
'����:      ����
'�ļ�:      modTooltips.bas
'====================================================

Option Explicit

Dim hWndTip         As Long                             '������ʾ�ı����ھ��

'����:      ����������ʾ�ı����ڣ��ڳ����ʼ��ʱ���ã�
'����ֵ:    �����Ĺ�����ʾ�ı����ھ��
Public Function CreateToolTip() As Long
    hWndTip = CreateWindowExA(0, "tooltips_class32", vbNullString, WS_POPUP Or TTS_ALWAYSTIP, _
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0, 0, App.hInstance, 0)
    
    SendMessageA hWndTip, TTM_SETTIPBKCOLOR, ByVal &H454242, 0                  '���ñ�����ɫ
    SendMessageA hWndTip, TTM_SETTIPTEXTCOLOR, ByVal &HF0F0F0, 0                '�����ı���ɫ
    SendMessageA hWndTip, TTM_SETMAXTIPWIDTH, 0, ByVal &HFFFFFFFF               '����Ϊ���е�
    
    CreateToolTip = hWndTip
End Function

'����:      �رչ�����ʾ�ı����ڣ��ڳ����˳�ʱ���ã�
Public Sub DestroyToolTip()
    DestroyWindow hWndTip
End Sub

'����:      Ϊָ���Ŀؼ���ӹ�����ʾ�ı�
'����:      TargetWindow: ��Ҫ��ӹ�����ʾ�ı��Ĵ���
'.          Tooltip: ������ʾ�ı�
'.          Title: ��ѡ�ģ�ָ��������ʾ�ı��ı���
'.          Icon: ��ѡ�ģ�ָ��������ʾ�ı���ͼ��
'����ֵ:    ����1��ʾ��ӳɹ���0��ʾ���ʧ��
Public Function CtlAddToolTip(TargetWindow As Long, Tooltip As String, _
    Optional Title As String, Optional Icon As Tooltip_Icon = 0) As Long
    
    Dim ti          As TTTOOLINFO
    Dim tmpStr()    As Byte
    
    tmpStr = StrconvEx(Tooltip)
    With ti
        .cbSize = Len(ti)
        .hWnd = TargetWindow
        .uFlags = TTF_IDISHWND Or TTF_SUBCLASS
        .uID = TargetWindow
        .lpszText = VarPtr(tmpStr(0))
    End With
    
    CtlAddToolTip = SendMessageA(hWndTip, TTM_ADDTOOL, 0, ByVal VarPtr(ti))
    
    If Len(Title) > 1 Then
        tmpStr = StrconvEx(Title)
        SendMessageA hWndTip, TTM_SETTITLE, ByVal Icon, ByVal VarPtr(tmpStr(0))
    End If
End Function


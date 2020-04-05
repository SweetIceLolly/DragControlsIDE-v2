Attribute VB_Name = "modGdbParser"
'====================================================
'����:      �ṩ����gdb����ĺ���
'����:      ����
'�ļ�:      modGdbParser.bas
'====================================================

Option Explicit

'������ö�ջ��Ϣ�ṹ
Public Type CallStackInfoStruct
    Address                 As String                                       '��ַ
    Args                    As String                                       '����
    File                    As String                                       '�ļ�
    Line                    As Long                                         '�к� (-1�����Ҫ���ļ��������ʾ���ļ�)
End Type

'����ģ����Ϣ�ṹ
Public Type ModuleInfoStruct
    File                    As String                                       'ģ���ļ�
    From                    As String                                       '�ӣ���ַ��
    To                      As String                                       '������ַ��
End Type

'�����߳���Ϣ�ṹ
Public Type ThreadInfoStruct
    Id                      As String                                       '�߳�ID
    Frame                   As String                                       '��ַ
End Type

'����:      ����gdb�Ķ�ջ���
'����:      strCallStack: ��Ҫ�����ĵ��ö�ջ���
'����ֵ:    �洢�ŵ��ö�ջ��Ϣ�Ľṹ
Public Function ParseCallStackString(strCallStack As String) As CallStackInfoStruct
    'On Error Resume Next
    Dim StrPos              As Long                                         '���ҵ����ַ�����λ��
    Dim Info                As CallStackInfoStruct
    
    If Mid(strCallStack, Len(strCallStack)) = vbCr Then                                 'ȥ���ַ�����β�Ļ��з�
        strCallStack = Left(strCallStack, Len(strCallStack) - 1)
    End If
    
    '��׼ȷ��ַ���ж�Ӧ�Ķ�̬�⺯�����ļ�
    If strCallStack Like "[#]* * in *(*)* from *:[\/]*" Then
        '����: #2  0x76926359 in KERNEL32!BaseThreadInitThunk () from C:\WINDOWS\SysWOW64\kernel32.dll
        Info.Line = -1                                                                      '���ΪҪ���ļ��������ʾ���ļ�
        StrPos = InStrRev(strCallStack, ":/")                                               '�����ַ����еġ�:/"���Էָ���ļ���
        If StrPos = 0 Then                                                                  '�Ҳ�����:/�����°�gdb���ͳ��Բ��ҡ�:\��
            StrPos = InStrRev(strCallStack, ":\")
        End If
        Info.File = Right(strCallStack, Len(strCallStack) - _
            InStrRev(strCallStack, " from ", StrPos) - 5)                                   '��#n  addr in func (args) from [file]��
        strCallStack = Left(strCallStack, Len(strCallStack) - Len(Info.File) - 6)           '��[#n  addr in func (args)] from file��
        StrPos = InStrRev(strCallStack, " (")                                               '�����ַ����еġ� (�����Էָ������
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '��#n  addr in func ([args])��
        StrPos = InStr(strCallStack, " 0x")                                                 '�����ַ����еġ� 0x�����Էָ����ַ
        Info.Address = Mid(strCallStack, StrPos + 1, _
            Len(strCallStack) - StrPos - Len(Info.Args) - 3)                                '��#n  [addr in func] (args)��
            
    '---------------------------------------------------------
    '��׼ȷ��ַ���ж�Ӧ�ļ�
    ElseIf strCallStack Like "[#]* * in *(*)* at *:[\/]*" Then
        '����: #1  0x0040144c in main () at C:\(aa) bb\.cpp:6
        StrPos = InStrRev(strCallStack, ":")                                                '�����ַ����еġ�:�����Էָ���к�
        Info.Line = CLng(Right(strCallStack, Len(strCallStack) - StrPos))                   '��#n  addr in func (args) at file:[line]��
        strCallStack = Left(strCallStack, StrPos - 1)                                       '��[#n  addr in func (args) at file]:line��
        StrPos = InStrRev(strCallStack, ":/")                                               '�����ַ����еġ�:/"���Էָ���ļ���
        If StrPos = 0 Then                                                                  '�Ҳ�����:/�����°�gdb���ͳ��Բ��ҡ�:\��
            StrPos = InStrRev(strCallStack, ":\")
        End If
        Info.File = Right(strCallStack, Len(strCallStack) - _
            InStrRev(strCallStack, " at ", StrPos) - 3)                                     '��#n  addr in func (args) at [file]��
        strCallStack = Left(strCallStack, Len(strCallStack) - Len(Info.File) - 4)           '��[#n  addr in func (args)] at file��
        StrPos = InStrRev(strCallStack, " (")                                               '�����ַ����еġ� (�����Էָ������
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '��#n  addr in func ([args])��
        StrPos = InStr(strCallStack, " 0x")                                                 '�����ַ����еġ� 0x�����Էָ����ַ
        Info.Address = Mid(strCallStack, StrPos + 1, _
            Len(strCallStack) - StrPos - Len(Info.Args) - 3)                                '��#n  [addr in func] (args)��
            
    '---------------------------------------------------------
    '��׼ȷ��ַ���޶�Ӧ�ļ�
    ElseIf strCallStack Like "[#]* * in *(*)*" Then
        '����: #1  0x00403c44 in main ()
        StrPos = InStrRev(strCallStack, " (")                                               '�����ַ����еġ� (�����Էָ������
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '��#n  addr in func ([args])��
        StrPos = InStr(strCallStack, " 0x")                                                 '�����ַ����еġ� 0x�����Էָ����ַ
        Info.Address = Mid(strCallStack, StrPos + 1, _
            Len(strCallStack) - StrPos - Len(Info.Args) - 3)                                '��#n  [addr in func] (args)��
    
    '---------------------------------------------------------
    '��׼ȷ��ַ���ж�Ӧ�ļ�
    ElseIf strCallStack Like "[#]* * (*)* at *:[\/]*" Then
        '����: #0  aaa (a=1, b=2, c=3, d=4) at C:\(aa) bb\.cpp:6
        StrPos = InStrRev(strCallStack, ":")                                                '�����ַ����еġ�:�����Էָ���к�
        Info.Line = CLng(Right(strCallStack, Len(strCallStack) - StrPos))                   '��#n  func (args) at file:[line]��
        strCallStack = Left(strCallStack, StrPos - 1)                                       '��[#n  func (args) at file]:line��
        StrPos = InStrRev(strCallStack, ":/")                                               '�����ַ����еġ�:/"���Էָ���ļ���
        If StrPos = 0 Then                                                                  '�Ҳ�����:/�����°�gdb���ͳ��Բ��ҡ�:\��
            StrPos = InStrRev(strCallStack, ":\")
        End If
        Info.File = Right(strCallStack, Len(strCallStack) - _
            InStrRev(strCallStack, " at ", StrPos) - 3)                                     '��#n  func (args) at [file]��
        strCallStack = Left(strCallStack, Len(strCallStack) - Len(Info.File) - 4)           '��[#n  func (args)] at file��
        StrPos = InStr(strCallStack, " ")                                                   '�����ַ����еġ� ������ȥ����ͷ�����
        strCallStack = Trim(Right(strCallStack, Len(strCallStack) - StrPos))                '��#n  [func (args)]��
        StrPos = InStrRev(strCallStack, " (")                                               '�����ַ����еġ� (�����Էָ������
        Info.Address = Left(strCallStack, StrPos - 1)                                       '��[func] (args)��
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '��func ([args])��
    
    '---------------------------------------------------------
    '��׼ȷ��ַ���޶�Ӧ�ļ�
    ElseIf strCallStack Like "[#]* * (*)*" Then
        '����: #1  aaa (a=1, b=2, c=3, d=4)
        StrPos = InStr(strCallStack, " ")                                                   '�����ַ����еġ� ������ȥ����ͷ�����
        strCallStack = Trim(Right(strCallStack, Len(strCallStack) - StrPos))                '��#n  [func (args)]��
        StrPos = InStrRev(strCallStack, " (")                                               '�����ַ����еġ� (�����Էָ������
        Info.Address = Left(strCallStack, StrPos - 1)                                       '��[func] (args)��
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '��func ([args])��
    
    '---------------------------------------------------------
    '����C++���У�����������ֱ����ӵ��б���
    Else
        StrPos = InStr(strCallStack, " ")                                                   '�����ַ����еġ� ������ȥ����ͷ�����
        Info.Address = Trim(Right(strCallStack, Len(strCallStack) - StrPos))                '��#n  [func (args)]��
    End If
    
    Info.File = Replace(Info.File, "/", "\")                                            '�ѵ�ַ��ġ�/���滻�ɡ�\��
    ParseCallStackString = Info
End Function

'����:      ����gdb��ģ�����
'����:      strModule: ��Ҫ������ģ�����
'����ֵ:    �洢��ģ����Ϣ�Ľṹ
Public Function ParseModuleString(strModule As String) As ModuleInfoStruct
    'on error resume next
    
    '���ӣ�
    '(gdb) info sharedlibrary
    'From        To          Syms Read   Shared Object Library
    '0x76920000  0x769e47b0  Yes (*)     C:\WINDOWS\SysWOW64\kernel32.dll
    '0x77051000  0x7724bfd0  Yes (*)     C:\WINDOWS\SysWOW64\KernelBase.dll
    '0x766c1000  0x7677e764  Yes (*)     C:\WINDOWS\SysWOW64\msvcrt.dll
    '(*): Shared library is missing debugging information.
    '(gdb)

    Dim StrPos              As Long                                     '���ҵ����ַ���λ��
    Dim Info                As ModuleInfoStruct
    
    If Mid(strModule, Len(strModule)) = vbCr Then                       'ȥ���ַ�����β�Ļ��з�
        strModule = Left(strModule, Len(strModule) - 1)
    End If
    
    '����ַ����Ƿ���ϸ�ʽ
    If strModule Like "*0x* 0x* * C:[\/]*" Then
        StrPos = InStr(strModule, "0x")                                     '������һ����0x������ȡ���ӡ���ַ
        Info.From = Mid(strModule, StrPos, 10)
        StrPos = InStr(StrPos + 10, strModule, "0x")                        '�����ڶ�����0x������ȡ��������ַ
        Info.To = Mid(strModule, StrPos, 10)
        StrPos = InStrRev(strModule, ":\")                                  '�ӽ�β��ǰ������:\����:/����
        If StrPos = 0 Then
            StrPos = InStrRev(strModule, ":/")
        End If
        StrPos = InStrRev(strModule, " ", StrPos)                           '���ҵ���λ����ǰ���ҿո񣬻�ȡ·��
        Info.File = Mid(strModule, StrPos + 1, Len(strModule) - StrPos)
    End If
    
    Info.File = Replace(Info.File, "/", "\")                            '�ѵ�ַ��ġ�/���滻�ɡ�\��
    ParseModuleString = Info
End Function

Public Function ParseThreadString(strThread As String) As ThreadInfoStruct
    'on error resume next
    
    '���ӣ�
    '(gdb) info threads
    '  Id   Target Id         Frame
    '  2    Thread 19152.0x17a0 0x77af3a4c in ?? ()
    '* 1    Thread 19152.0x4794 main () at C:\Users\12574\Documents\MyProjects\te\te.cpp:2
    '(gdb)
    
    Dim StrPos              As Long                                     '���ҵ����ַ���λ��
    Dim StrPos2             As Long
    Dim Info                As ThreadInfoStruct
    
    If Mid(strThread, Len(strThread)) = vbCr Then                       'ȥ���ַ�����β�Ļ��з�
        strThread = Left(strThread, Len(strThread) - 1)
    End If
    
    '����ַ����Ƿ���ϸ�ʽ
    If strThread Like "* * *.0x* *" Then
        StrPos = InStr(strThread, ".0x")
        StrPos = InStr(StrPos, strThread, " ")                              '�ӡ�.0x����������� ��
        Info.Frame = Right(strThread, Len(strThread) - StrPos)              '��ȡ�ո���������Ϊ��ַ
        StrPos2 = InStrRev(strThread, " ", StrPos - 1) + 1                  '�ӡ�.0x����ǰ������ ��
        Info.Id = Mid(strThread, StrPos2, StrPos - StrPos2)                 '��ȡ��.0x��ǰ��Ŀո�ͺ���Ŀո��м���ı���ΪID
    End If
    
    ParseThreadString = Info
End Function

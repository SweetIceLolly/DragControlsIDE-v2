Attribute VB_Name = "modGdbParser"
Option Explicit

'������ö�ջ��Ϣ�ṹ
Public Type CallStackInfoStruct
    Address                                     As String                                   '��ַ
    Args                                        As String                                   '����
    File                                        As String                                   '�ļ�
    Line                                        As Long                                     '�к�
End Type

'����:      ����gdb�Ķ�ջ���
'����:      strCallStack: ��Ҫ�����ĵ��ö�ջ���
'����ֵ:    �洢�ŵ��ö�ջ��Ϣ�Ľṹ
Public Function ParseCallStackString(strCallStack As String) As CallStackInfoStruct
    'On Error Resume Next
    
    Dim StrPos              As Long                                         '���ҵ����ַ�����λ��
    Dim BracketLevel        As Long                                         '����ƥ�������һ��ʼ��0��������(����1, ������)����1
    Dim Info                As CallStackInfoStruct
    
    If strCallStack Like "[#]* * in *(*) at *:\*" Then                                  '��׼ȷ��ַ���ж�Ӧ�ļ�
        '����: #1  0x0040144c in main () at C:\(aa) bb\.cpp:6
        StrPos = InStrRev(strCallStack, ":")                                                '�����ַ����еġ�:�����Էָ���к�
        Info.Line = CLng(Right(strCallStack, Len(strCallStack) - StrPos))                   '��#n  addr in func (args) at file:[line]��
        strCallStack = Left(strCallStack, StrPos - 1)                                       '��[#n  addr in func (args) at file]:line��
        StrPos = InStrRev(strCallStack, ":\")                                               '�����ַ����еġ�:\"���Էָ���ļ���
        If StrPos = 0 Then                                                                  '�Ҳ�����:\�����°�gdb���ͳ��Բ��ҡ�:/�����ɰ�gdb��
            StrPos = InStrRev(strCallStack, ":/")
        End If
        Info.File = Right(strCallStack, Len(strCallStack) - _
            InStrRev(strCallStack, " at ", StrPos) - 3)                                     '��#n  addr in func (args) at [file]��
        strCallStack = Left(strCallStack, Len(strCallStack) - Len(Info.File) - 4)           '��[#n  addr in func (args)] at file��
        StrPos = InStr(strCallStack, " (")                                                  '�����ַ����еġ� (�����Էָ������
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '��#n  addr in func ([args])��
        StrPos = InStr(strCallStack, " 0x")                                                 '�����ַ����еġ� 0x�����Էָ����ַ
        Info.Address = Mid(strCallStack, StrPos + 1, _
            Len(strCallStack) - StrPos - Len(Info.Args) - 3)                                '��#n  [addr in func] (args)��
    ElseIf strCallStack Like "[#]* * in *(*)" Then                                      '��׼ȷ��ַ���޶�Ӧ�ļ�
        '����: #1  0x00403c44 in main ()
        StrPos = InStr(strCallStack, " (")                                                  '�����ַ����еġ� (�����Էָ������
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '��#n  addr in func ([args])��
        StrPos = InStr(strCallStack, " 0x")                                                 '�����ַ����еġ� 0x�����Էָ����ַ
        Info.Address = Mid(strCallStack, StrPos + 1, _
            Len(strCallStack) - StrPos - Len(Info.Args) - 3)                                '��#n  [addr in func] (args)��
    ElseIf strCallStack Like "[#]* * (*) at *:\*" Then                                  '��׼ȷ��ַ���ж�Ӧ�ļ�
        '����: #0  aaa (a=1, b=2, c=3, d=4) at C:\(aa) bb\.cpp:6
        StrPos = InStrRev(strCallStack, ":")                                                '�����ַ����еġ�:�����Էָ���к�
        Info.Line = CLng(Right(strCallStack, Len(strCallStack) - StrPos))                   '��#n  func (args) at file:[line]��
        strCallStack = Left(strCallStack, StrPos - 1)                                       '��[#n  func (args) at file]:line��
        StrPos = InStrRev(strCallStack, ":\")                                               '�����ַ����еġ�:\"���Էָ���ļ���
        If StrPos = 0 Then                                                                  '�Ҳ�����:\�����°�gdb���ͳ��Բ��ҡ�:/�����ɰ�gdb��
            StrPos = InStrRev(strCallStack, ":/")
        End If
        Info.File = Right(strCallStack, Len(strCallStack) - _
            InStrRev(strCallStack, " at ", StrPos) - 3)                                     '��#n  func (args) at [file]��
        strCallStack = Left(strCallStack, Len(strCallStack) - Len(Info.File) - 4)       '��[#n  func (args)] at file��
        StrPos = InStr(strCallStack, " ")                                                   '�����ַ����еġ� ������ȥ����ͷ�����
        strCallStack = Trim(Right(strCallStack, Len(strCallStack) - StrPos))                '��#n  [func (args)]��
        StrPos = InStr(strCallStack, " (")                                                  '�����ַ����еġ� (�����Էָ������
        Info.Address = Left(strCallStack, StrPos - 1)                                       '��[func] (args)��
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '��func ([args])��
    ElseIf strCallStack Like "[#]* * (*)" Then                                          '��׼ȷ��ַ���޶�Ӧ�ļ�
        '����: #1  aaa (a=1, b=2, c=3, d=4)
        StrPos = InStr(strCallStack, " ")                                                   '�����ַ����еġ� ������ȥ����ͷ�����
        strCallStack = Trim(Right(strCallStack, Len(strCallStack) - StrPos))                '��#n  [func (args)]��
        StrPos = InStr(strCallStack, " (")                                                  '�����ַ����еġ� (�����Էָ������
        Info.Address = Left(strCallStack, StrPos - 1)                                       '��[func] (args)��
        Info.Args = Mid(strCallStack, StrPos + 2, Len(strCallStack) - StrPos - 2)           '��func ([args])��
    End If
    
    ParseCallStackString = Info
End Function



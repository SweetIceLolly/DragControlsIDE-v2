VERSION 5.00
Begin VB.Form frmCallStack 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "���ö�ջ"
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DragControlsIDE.DarkListView lvCallStack 
      Height          =   2655
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4683
   End
End
Attribute VB_Name = "frmCallStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'����:      ���ö�ջ���ڣ����ж�״̬����ʾ���ö�ջ
'����:      ����
'�ļ�:      frmCallStack.frm
'====================================================

Option Explicit

'������ö�ջ��Ϣ�ṹ
Private Type CallStackInfoStruct
    Address                 As String                                       '��ַ
    Args                    As String                                       '����
    File                    As String                                       '�ļ�
    Line                    As Long                                         '�к�
End Type

Dim CallStackInfo()         As CallStackInfoStruct                          '���е��ö�ջ��Ϣ

'����:      ������ж�����Ϊ��һ�ε�����׼��
Public Sub ClearEverything()
    Me.lvCallStack.Clear
    ReDim CallStackInfo(0)
End Sub

'����:      ��ȡ���ö�ջ�б�
Public Sub GetCallStack()
    'On Error Resume Next       'todo
    Dim PipeOutput          As String                                       '�ܵ������
    Dim OutputLines()       As String                                       '�����ÿһ��
    Dim StrPos              As Long                                         '���ҵ����ַ�����λ��
    Dim BracketLevel        As Long                                         '����ƥ�������һ��ʼ��0��������(����1, ������)����1
    Dim NewListItem         As Long                                         '����ӵ�ListView�б�������
    Dim i                   As Long
    
    Me.lvCallStack.Clear
    frmMain.DockingPane.Panes(10).Title = Lang_CallStack_Retrieving_Caption
    
    frmMain.GdbPipe.ClearPipe                                                                               '��չܵ��������
    frmMain.GdbPipe.DosInput "info stack" & vbCrLf                                                          '��gdb���ͻ�ȡ���ö�ջ����
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                                                          '��ȡgdb���
    
    OutputLines = Split(PipeOutput, vbCrLf)                                                                 '���зָ���
    ReDim CallStackInfo(UBound(OutputLines) - 1)                                                            '������Ϣ�б�Ԫ��
    For i = 0 To UBound(OutputLines)                                                                        '���н��з���
        If Trim(OutputLines(i)) <> "(gdb)" Then                                                                 'ȥ�����������(gdb) ��
            If Mid(OutputLines(i), Len(OutputLines(i))) = vbCr Then                                                 'ȥ���ַ�����β�Ļ��з�
                OutputLines(i) = Left(OutputLines(i), Len(OutputLines(i)) - 1)
            End If
            If OutputLines(i) Like "[#]* * in *(*) at *:\*" Then                                                    '��׼ȷ��ַ���ж�Ӧ�ļ�
                '����: #1  0x0040144c in main () at C:\(aa) bb\.cpp:6
                StrPos = InStrRev(OutputLines(i), ":")                                                                  '�����ַ����еġ�:�����Էָ���к�
                CallStackInfo(i).Line = CLng(Right(OutputLines(i), Len(OutputLines(i)) - StrPos))                       '��#n  addr in func (args) at file:[line]��
                OutputLines(i) = Left(OutputLines(i), StrPos - 1)                                                       '��[#n  addr in func (args) at file]:line��
                StrPos = InStrRev(OutputLines(i), ":\")                                                                 '�����ַ����еġ�:\"���Էָ���ļ���
                If StrPos = 0 Then                                                                                      '�Ҳ�����:\�����°�gdb���ͳ��Բ��ҡ�:/�����ɰ�gdb��
                    StrPos = InStrRev(OutputLines(i), ":/")
                End If
                CallStackInfo(i).File = Right(OutputLines(i), Len(OutputLines(i)) - _
                    InStrRev(OutputLines(i), " at ", StrPos) - 3)                                                       '��#n  addr in func (args) at [file]��
                OutputLines(i) = Left(OutputLines(i), Len(OutputLines(i)) - Len(CallStackInfo(i).File) - 4)             '��[#n  addr in func (args)] at file��
                StrPos = InStr(OutputLines(i), " (")                                                                    '�����ַ����еġ� (�����Էָ������
                CallStackInfo(i).Args = Mid(OutputLines(i), StrPos + 2, Len(OutputLines(i)) - StrPos - 2)               '��#n  addr in func ([args])��
                StrPos = InStr(OutputLines(i), " 0x")                                                                   '�����ַ����еġ� 0x�����Էָ����ַ
                CallStackInfo(i).Address = Mid(OutputLines(i), StrPos + 1, _
                    Len(OutputLines(i)) - StrPos - Len(CallStackInfo(i).Args) - 3)                                      '��#n  [addr in func] (args)��
            ElseIf OutputLines(i) Like "[#]* * in *(*)" Then                                                        '��׼ȷ��ַ���޶�Ӧ�ļ�
                '����: #1  0x00403c44 in main ()
                StrPos = InStr(OutputLines(i), " (")                                                                    '�����ַ����еġ� (�����Էָ������
                CallStackInfo(i).Args = Mid(OutputLines(i), StrPos + 2, Len(OutputLines(i)) - StrPos - 2)               '��#n  addr in func ([args])��
                StrPos = InStr(OutputLines(i), " 0x")                                                                   '�����ַ����еġ� 0x�����Էָ����ַ
                CallStackInfo(i).Address = Mid(OutputLines(i), StrPos + 1, _
                    Len(OutputLines(i)) - StrPos - Len(CallStackInfo(i).Args) - 3)                                      '��#n  [addr in func] (args)��
            ElseIf OutputLines(i) Like "[#]* * (*) at *:\*" Then                                                    '��׼ȷ��ַ���ж�Ӧ�ļ�
                '����: #0  aaa (a=1, b=2, c=3, d=4) at C:\(aa) bb\.cpp:6
                StrPos = InStrRev(OutputLines(i), ":")                                                                  '�����ַ����еġ�:�����Էָ���к�
                CallStackInfo(i).Line = CLng(Right(OutputLines(i), Len(OutputLines(i)) - StrPos))                       '��#n  func (args) at file:[line]��
                OutputLines(i) = Left(OutputLines(i), StrPos - 1)                                                       '��[#n  func (args) at file]:line��
                StrPos = InStrRev(OutputLines(i), ":\")                                                                 '�����ַ����еġ�:\"���Էָ���ļ���
                If StrPos = 0 Then                                                                                      '�Ҳ�����:\�����°�gdb���ͳ��Բ��ҡ�:/�����ɰ�gdb��
                    StrPos = InStrRev(OutputLines(i), ":/")
                End If
                CallStackInfo(i).File = Right(OutputLines(i), Len(OutputLines(i)) - _
                    InStrRev(OutputLines(i), " at ", StrPos) - 3)                                                       '��#n  func (args) at [file]��
                OutputLines(i) = Left(OutputLines(i), Len(OutputLines(i)) - Len(CallStackInfo(i).File) - 4)             '��[#n  func (args)] at file��
                StrPos = InStr(OutputLines(i), " ")                                                                     '�����ַ����еġ� ������ȥ����ͷ�����
                OutputLines(i) = Trim(Right(OutputLines(i), Len(OutputLines(i)) - StrPos))                              '��#n  [func (args)]��
                StrPos = InStr(OutputLines(i), " (")                                                                    '�����ַ����еġ� (�����Էָ������
                CallStackInfo(i).Address = Left(OutputLines(i), StrPos - 1)                                             '��[func] (args)��
                CallStackInfo(i).Args = Mid(OutputLines(i), StrPos + 2, Len(OutputLines(i)) - StrPos - 2)               '��func ([args])��
            ElseIf OutputLines(i) Like "[#]* * (*)" Then                                                            '��׼ȷ��ַ���޶�Ӧ�ļ�
                '����: #1  aaa (a=1, b=2, c=3, d=4)
                StrPos = InStr(OutputLines(i), " ")                                                                     '�����ַ����еġ� ������ȥ����ͷ�����
                OutputLines(i) = Trim(Right(OutputLines(i), Len(OutputLines(i)) - StrPos))                              '��#n  [func (args)]��
                StrPos = InStr(OutputLines(i), " (")                                                                    '�����ַ����еġ� (�����Էָ������
                CallStackInfo(i).Address = Left(OutputLines(i), StrPos - 1)                                             '��[func] (args)��
                CallStackInfo(i).Args = Mid(OutputLines(i), StrPos + 2, Len(OutputLines(i)) - StrPos - 2)               '��func ([args])��
            End If
            
            NewListItem = Me.lvCallStack.AddItem(CallStackInfo(i).Address)                                          '������б���
            If CallStackInfo(i).Args <> "" Then
                Me.lvCallStack.SetItemText CallStackInfo(i).Args, NewListItem, 1
            Else
                CallStackInfo(i).Args = Lang_CallStack_NoArg
                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 1
            End If
            If CallStackInfo(i).File <> "" Then
                Me.lvCallStack.SetItemText CallStackInfo(i).File, NewListItem, 2
            Else
                CallStackInfo(i).File = Lang_CallStack_NoArg
                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 2
            End If
            If CallStackInfo(i).Line <> 0 Then
                Me.lvCallStack.SetItemText CStr(CallStackInfo(i).Line), NewListItem, 3
            Else
                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 3
            End If
        
'            If OutputLines(i) Like "[#]* * in *(*)*" Then                                                           '����в����ļ���
'
'                OutputLines(i) = Right(OutputLines(i), Len(OutputLines(i)) - Len(Split(OutputLines(i), " ")(0)) - 1)    '��#n func(arg types) (args)��
'                CallStackInfo(i).Address = OutputLines(i)
'                StrPos = InStr(OutputLines(i), "(")                                                                     '�����ַ������һ����(��
'                CallStackInfo(i).Args = Mid(OutputLines(i), StrPos + 1, Len(OutputLines(i)) - StrPos - 2)               '��#n func(arg types) ([args])��
'                If CallStackInfo(i).Args = "" Then
'                    CallStackInfo(i).Args = Lang_CallStack_NoArg
'                End If
'
'                NewListItem = Me.lvCallStack.AddItem(CallStackInfo(i).Address)                                          '������б���
'                Me.lvCallStack.SetItemText "", NewListItem, 1
'                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 2
'                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 3
'            Else                                                                                                    '����д����ļ���
'                StrPos = InStrRev(OutputLines(i), ":")                                                                  '��#n func(arg types) (args) at file:line��
'                CallStackInfo(i).Line = CLng(Right(OutputLines(i), Len(OutputLines(i)) - StrPos))                       '��#n func(arg types) (args) at file:[line]��
'                OutputLines(i) = Left(OutputLines(i), StrPos - 1)                                                       '��[#n func(arg types) (args) at file]:line��
'                StrPos = InStrRev(OutputLines(i), ":\")                                                                 '��ǰ���ҡ�:\�����°�gdb��
'                If StrPos = 0 Then                                                                                      '�Ҳ�����:\���Ͳ��ҡ�:/�����ɰ�gdb��
'                    StrPos = InStrRev(OutputLines(i), ":/")
'                End If
'                StrPos = InStrRev(OutputLines(i), " at ", StrPos)                                                       '�ӡ�:/����λ�ü�����ǰ���ҡ� at ��
'                CallStackInfo(i).File = Replace(Right(OutputLines(i), Len(OutputLines(i)) - StrPos - 3), "/", "\")      '��#n func(arg types) (args) at [file]��
'                OutputLines(i) = Left(OutputLines(i), StrPos - 1)                                                       '��[#n func(arg types) (args)] at file��
'                OutputLines(i) = Right(OutputLines(i), Len(OutputLines(i)) - InStr(OutputLines(i), " ") - 1)            '��#n [func(arg types) (args)]��
'                StrPos = InStr(OutputLines(i), "(")                                                                     '�����ַ�����ĵ�һ����(��
'                BracketLevel = 0
'                For StrPos = StrPos + 1 To Len(OutputLines(i))                                                          '���������ƥ��ġ�)�����ⲿ�ִ�����frmLocals��ArrayParser�еĴ������ƣ�
'                    If Mid(OutputLines(i), StrPos, 1) = "(" Then                                                            '����������: ����+1
'                        BracketLevel = BracketLevel + 1
'                    ElseIf Mid(OutputLines(i), StrPos, 1) = ")" Then                                                        '����������
'                        If BracketLevel <= 0 Then                                                                               '���ż���Ϊ0���������Ѿ�ƥ�䡣��ʱStrPos����һ��ƥ��ġ�)����λ��
'                            CallStackInfo(i).Address = Left(OutputLines(i), StrPos)                                                 '��[func(arg types)] (args)��
'                            Exit For                                                                                                '�������������
'                        Else                                                                                                    '������δƥ�䣬������1�������������
'                            BracketLevel = BracketLevel - 1
'                        End If
'                    ElseIf Mid(OutputLines(i), StrPos, 1) = """" Then                                                       '������"�������ҵ���һ��ƥ��ġ�"����ȷ������������ַ����м�ȥ
'                        Do                                                                                                      'һֱ�����ҡ�"����ֱ���������ַ����м�
'                            StrPos = StrPos + 1
'                        Loop Until (Mid(OutputLines(i), StrPos, 1) = """" And Mid(OutputLines(i), StrPos - 1, 1) <> "\") Or StrPos > Len(OutputLines(i))
'                    End If
'                Next StrPos
'                If StrPos = Len(OutputLines(i)) Then                                                                    '�������û�в���
'                    CallStackInfo(i).Args = Lang_CallStack_NoArg                                                            '���ò���Ϊ��
'                Else                                                                                                    '��������в���
'                    CallStackInfo(i).Args = Right(OutputLines(i), Len(OutputLines(i)) - StrPos - 1)                         '��func(arg types) [(args)]��
'                End If
'
'                NewListItem = Me.lvCallStack.AddItem(CallStackInfo(i).Address)                                          '������б���
'                Me.lvCallStack.SetItemText CallStackInfo(i).Args, NewListItem, 1
'                Me.lvCallStack.SetItemText CallStackInfo(i).File, NewListItem, 2
'                Me.lvCallStack.SetItemText CStr(CallStackInfo(i).Line), NewListItem, 3
'            End If
        End If
    Next i
    
    frmMain.DockingPane.Panes(10).Title = Lang_CallStack_Caption
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_CallStack_Caption
    
    Me.lvCallStack.Move 0, 0
    
    Me.lvCallStack.AddColumnHeader Lang_Breakpoints_ListViewHeader_Address, 300
    Me.lvCallStack.AddColumnHeader Lang_CallStack_Args, 300
    Me.lvCallStack.AddColumnHeader Lang_Breakpoints_ListViewHeader_File, 150
    Me.lvCallStack.AddColumnHeader Lang_Breakpoints_ListViewHeader_Line
    
    ReDim CallStackInfo(0)                                                  '��ʼ�����ö�ջ��Ϣ�б�
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.lvCallStack.Width = Me.ScaleWidth
    Me.lvCallStack.Height = Me.ScaleHeight
End Sub

Private Sub lvCallStack_Click(iItem As Long, iSubItem As Long, X As Long, Y As Long)
    'On Error Resume Next       'todo
    
    CtlAddToolTip Me.lvCallStack.ListViewHwnd, Lang_Breakpoints_ListViewHeader_Address & ": " & CallStackInfo(iItem).Address & vbCrLf & _
        Lang_CallStack_Args & ": " & CallStackInfo(iItem).Args & vbCrLf & _
        Lang_Breakpoints_ListViewHeader_File & ": " & CallStackInfo(iItem).File & ":" & CallStackInfo(iItem).Line, _
        Lang_CallStack_Tooltip_Title, TTI_INFO
End Sub

Private Sub lvCallStack_DoubleClick(iItem As Long, iSubItem As Long, X As Long, Y As Long)
    On Error Resume Next
    Dim i                   As Long
    
    If CallStackInfo(iItem).File <> "" Then                                                 '����ж�Ӧ���ļ�
        For i = 0 To UBound(CurrentProject.Files)                                               '�����ڹ��̵��ļ��в��Ҷ�Ӧ���ļ�
            If CurrentProject.Files(i).FilePath = CallStackInfo(iItem).File Then                    '���ҵ���Ӧ���ļ�
                If CurrentProject.Files(i).TargetWindow Is Nothing Then                              '����ж�Ӧ�Ĵ��봰�ھ��л���ȥ
                    Dim NewCodeWindow   As frmCodeWindow
                    Dim FileData        As String
                    Dim tmpData         As String
                    
                    Set NewCodeWindow = CreateNewCodeWindow(i)                                              '�����µĴ��봰�岢���ð󶨵��ļ����
                    NewCodeWindow.Caption = GetFileName(CallStackInfo(iItem).File)
                    
                    Err.Clear
                    Open CallStackInfo(iItem).File For Input As #1                                          '���Դ򿪶�Ӧ�Ĵ����ļ�
                        If Err.Number <> 0 Then
                            Close #1
                            NoSkinMsgBox Lang_Main_Debug_OpenSourceFailure & CallStackInfo(iItem).File, vbExclamation, Lang_Msgbox_Error
                        Else
                            Do While Not EOF(1)
                                Line Input #1, tmpData
                                FileData = FileData & tmpData & vbCrLf
                            Loop
                        End If
                    Close #1
                    
                    frmMain.TabBar.AddForm NewCodeWindow
                Else                                                                                    'û�еĻ��ʹ���һ���µĴ��봰��
                    frmMain.TabBar.SwitchToByForm CurrentProject.Files(i).TargetWindow
                End If
                
                CurrentProject.Files(i).TargetWindow.SyntaxEdit.CurrPos.Row = CallStackInfo(iItem).Line
                CurrentProject.Files(i).TargetWindow.SyntaxEdit.SetFocus
                Exit Sub
            End If
        Next i
    End If
End Sub


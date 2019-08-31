VERSION 5.00
Begin VB.Form frmLocals 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "����"
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSelMargin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00373333&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin DragControlsIDE.DarkListView lvLocals 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4683
   End
   Begin VB.Image imgExpanded 
      Height          =   240
      Left            =   7200
      Picture         =   "frmLocals.frx":0000
      Top             =   3600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFolded 
      Height          =   240
      Left            =   6600
      Picture         =   "frmLocals.frx":038A
      Top             =   3600
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmLocals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'����:      ���ر������ڣ���ʾ��ǰ������ı������ơ�ֵ�����ͣ��������û��Ա�����ֵ�����޸�
'����:      ����
'�ļ�:      frmLocals.frm
'====================================================

Option Explicit

'���������Ϣ�ṹ
Private Type NodeInfo
    ParentNode                  As Long                                     'ĸ�ڵ����ţ�-1������㣩
    ChildNodes()                As Long                                     '�����ӽڵ����ţ�����-1��ֹͣ��
    VarName                     As String                                   '������
    Value                       As String                                   'ֵ
    TypeName                    As String                                   '������
    ListViewItemIndex           As Long                                     '��ListView�����Ӧ���б��-1����û�ж�Ӧ���б��
    Expanded                    As Boolean                                  '�ýڵ��Ƿ���չ����True = ��չ����
End Type

Dim VarNodes()                  As NodeInfo                                 '���б�����Ϣ�����һ��Ԫ���Ƕ���ģ�
Dim ColumnHeaderHeight          As Long                                     'ListView��ColumnHeader�߶�
Dim ListItemHeight              As Long                                     'ListViewÿ���б���ĸ߶�
Dim SpaceCount                  As Integer                                  'ͼƬ��Ŀ���൱�ڶ��ٸ��ո�
Dim ColumnHeader                As Long                                     '�б�ͷ�Ĵ��ھ��

'����:      ������ж�����Ϊ��һ�ε�����׼��
Public Sub ClearEverything()
    Me.lvLocals.Clear
    ReDim VarNodes(0)
    Me.picSelMargin.Cls
End Sub

'����:      �۵�ָ���Ľڵ㼰���ӽڵ�
'����:      VarNodeIndex: ��Ҫ���۵��Ľڵ����
Private Sub FoldItem(VarNodeIndex As Long)
    'On Error Resume Next       'todo
    Dim i                       As Long
    Dim j                       As Long
    
    For i = UBound(VarNodes(VarNodeIndex).ChildNodes) - 1 To 0 Step -1      '�����һ����ӽڵ㣨��������
        If VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).Expanded Then         '�����һ���ӽڵ㻹���ӽڵ㣬����ɾ�������ӽڵ㣬�ٱ�ǳ����۵�
            Call FoldItem(VarNodes(VarNodeIndex).ChildNodes(i))
            VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).Expanded = False
        End If
        
        'ɾ�����б���
        Me.lvLocals.DeleteItem VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).ListViewItemIndex
        VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).ListViewItemIndex = -1
        VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).Expanded = False
    Next i
    
    '��ɾ�б���֮����б�����ܵ�ɾ��Ӱ����б����ǰ��
    For j = VarNodeIndex To UBound(VarNodes) - 1
        If VarNodes(j).ListViewItemIndex > VarNodes(VarNodeIndex).ListViewItemIndex Then
            VarNodes(j).ListViewItemIndex = VarNodes(j).ListViewItemIndex - UBound(VarNodes(VarNodeIndex).ChildNodes)
        End If
    Next j
End Sub

'����:      չ��ָ���Ľڵ�
'����:      VarNodeIndex: ��Ҫ���۵��Ľڵ����
Private Sub ExpandItem(VarNodeIndex As Long)
    'On Error Resume Next       'todo
    Dim Level                   As Long                                     '��ǰ�ڵ㴦�ڵڼ���
    Dim ParentIndex             As Long                                     '��ǰ�ڵ�����Ӧ��ĸ�ڵ������
    Dim NewItemIndex            As Long                                     '����ӵ�ListView�б��������
    Dim i                       As Long
    
    ParentIndex = VarNodeIndex
    Level = 1
    Do                                                                      '����ӵ�ǰ�ڵ㵽���Ĳ���
        ParentIndex = VarNodes(ParentIndex).ParentNode
        Level = Level + 1
    Loop Until ParentIndex = -1
    
    NewItemIndex = VarNodes(VarNodeIndex).ListViewItemIndex                 '�����б�����뵽��ǰ�ڵ�ĺ���
    For i = VarNodeIndex + 1 To UBound(VarNodes) - 1                        '����VarNodes��������Ӧ���б����������б���֮��ľͰ��������
        If VarNodes(i).ListViewItemIndex > NewItemIndex Then
            VarNodes(i).ListViewItemIndex = VarNodes(i).ListViewItemIndex + UBound(VarNodes(VarNodeIndex).ChildNodes)
        End If
    Next i
    For i = 0 To UBound(VarNodes(VarNodeIndex).ChildNodes) - 1              '���������һ���ӽڵ���б���
        NewItemIndex = Me.lvLocals.AddItem(Space(SpaceCount * Level) & VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).VarName, NewItemIndex + 1)
        VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).ListViewItemIndex = NewItemIndex
        Me.lvLocals.SetItemText VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).TypeName, NewItemIndex, 1
        Me.lvLocals.SetItemText VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).Value, NewItemIndex, 2
    Next i
End Sub

'����:      �ػ����еĽڵ�ͼ��
Public Sub RedrawNodeIcons()
    On Error Resume Next
    Dim i                       As Long
    Dim TopItem                 As Long, BottomItem                 As Long

    Me.picSelMargin.Cls
    TopItem = Me.lvLocals.GetTopIndex()                                     '��ȡListView�е�һ�����ӵ��б�������
    BottomItem = TopItem + Me.lvLocals.Height / ListItemHeight              '�����ListView�����һ�����ӵ��б�������
    For i = 0 To UBound(VarNodes) - 1
         If VarNodes(i).ChildNodes(0) <> -1 Then                                 '�������������ӽڵ�
             If VarNodes(i).ListViewItemIndex >= TopItem And _
                VarNodes(i).ListViewItemIndex <= BottomItem Then                     '�����ڿ��ӵķ�Χ��

                 Me.picSelMargin.PaintPicture IIf(VarNodes(i).Expanded, Me.imgExpanded.Picture, Me.imgFolded.Picture), _
                     0, (VarNodes(i).ListViewItemIndex - TopItem) * ListItemHeight + 60
             End If
         End If
    Next i
End Sub

'����:      ��VarNodes��ListView���������Ŀ
'����:      ParentNode: �������㣬��Ϊ-1; ����Ϊĸ�ڵ�����
'.          VarName: ʵ�ʵı�����
'.          DisplayName: ��ʾ������
'.          TypeName: ����������
'.          Value: ������ֵ
'����ֵ:    ����ӵ�VarNodesԪ������
Private Function AddVarItem(ParentNode As Long, VarName As String, DisplayName As String, TypeName As String, Value As String) As Long
    'On Error Resume Next       'todo
    Dim NewItemIndex            As Long
    
    NewItemIndex = UBound(VarNodes)
    With VarNodes(NewItemIndex)                                             '���ñ�����Ϣ
        .ParentNode = ParentNode
        .TypeName = TypeName
        .Value = Value
        .VarName = VarName
        ReDim .ChildNodes(0)                                                    '��ʼ���ӽڵ�����
        .ChildNodes(0) = -1                                                     '-1��ʾû���ӽڵ�
    End With
    AddVarItem = NewItemIndex
    ReDim Preserve VarNodes(NewItemIndex + 1)                               '���������Ϣ����
    If ParentNode <> -1 Then                                                '����µ���Ŀ���ڵ�һ��
        With VarNodes(ParentNode)
            .ChildNodes(UBound(.ChildNodes)) = NewItemIndex                         '�������Ŀ��ĸ�ڵ���ӽڵ����飬���Ѹ���Ŀ����ż�¼��ȥ
            ReDim Preserve .ChildNodes(UBound(.ChildNodes) + 1)
            .ChildNodes(UBound(.ChildNodes)) = -1                                   '���������Ԫ������Ϊ-1������û���ӽڵ�
        End With
        VarNodes(NewItemIndex).ListViewItemIndex = -1                           '�Ѷ�Ӧ���б�����Ҫ����Ϊ-1������û�ж�Ӧ���б���
    Else                                                                    '����µ���Ŀ�ڵ�һ�㣬�Ͱ�����ӵ�ListView��
        '����б�������ı�
         VarNodes(NewItemIndex).ListViewItemIndex = Me.lvLocals.AddItem(Space(SpaceCount) & DisplayName)
         NewItemIndex = VarNodes(NewItemIndex).ListViewItemIndex
         Me.lvLocals.SetItemText TypeName, NewItemIndex, 1
         Me.lvLocals.SetItemText Value, NewItemIndex, 2
    End If
End Function

'����:      �������硰x = 0x12345 "*"�����ַ����������
'����:      ParentItem: ĸ�ڵ����
'.          OutputString: ��Ҫ�������ַ���
'.          NewVarNodesIndex: ��������ӵ�VarNodesԪ�����
'����ֵ:    ��������ĳ���
Private Function StringParser(ParentItem As Long, OutputString As String, Optional ByRef NewVarNodesIndex As Long) As Long
    'On Error Resume Next       'todo
    Dim VarName                 As String                                   '�������������ƣ�������a.b.c��
    Dim VarTypeName             As String                                   '��������
    Dim VarValue                As String                                   '������ֵ
    Dim PipeOutput              As String                                   '�ܵ������
    Dim SplitTmp()              As String                                   '�ַ����ԡ� = �����зָ�Ļ���
    
    VarName = Split(OutputString, " = ")(0)                                 '�ָ��������
    
    '��ȡ������ֵ
    frmMain.GdbPipe.ClearPipe                                               '��չܵ��������
    frmMain.GdbPipe.DosInput "p (" & VarName & ")" & vbCrLf                 '��gdb���ͼ�����ʽ���p (var)��
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                          '��ȡgdb���
    PipeOutput = Split(PipeOutput, vbCrLf)(0)                               'ֻ��Ҫ����ĵ�һ��
    SplitTmp = Split(PipeOutput, " = ")                                     '��[var] = [value]��
    If UBound(SplitTmp) > 0 Then                                            '�������������ڡ�* = *��
        VarValue = Right(PipeOutput, Len(PipeOutput) - Len(SplitTmp(0)) - 3)    '��var = [value]��
    Else                                                                    '��Ѱ�������
        VarValue = Lang_Locals_Error
    End If
        
    '��ȡ��������
    frmMain.GdbPipe.ClearPipe                                               '��չܵ��������
    frmMain.GdbPipe.DosInput "ptype (" & VarName & ")" & vbCrLf             '��gdb���ͻ�ȡ���ʽ�������ptype (var)��
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                          '��ȡgdb���
    PipeOutput = Left(PipeOutput, InStrRev(PipeOutput, vbCrLf))             'ȥ����������һ��
    If Left(PipeOutput, 7) = "type = " Then                                 '��������ڡ�type = *��
        VarTypeName = Right(PipeOutput, Len(PipeOutput) - 7)                    '��type = [*]��
    Else
        VarTypeName = Lang_Locals_Error
    End If
    
    SplitTmp = Split(VarName, ".")                                          '�ԡ�.���ָ��������[a].[b].[c]����Ȼ��ֻ�������һ�����ƣ�a.b.[c]��
    NewVarNodesIndex = AddVarItem(ParentItem, VarName, SplitTmp(UBound(SplitTmp)), VarTypeName, VarValue)
                
    If Left(VarValue, 1) = "{" And ParentItem <> -1 Then                    '���������������ֵ�ԡ�{����ͷ�����ҷǵ�һ�������˵���ܽ�һ������
        If Left(VarValue, 2) = "{{" Then                                        '����������ǡ�{{����˵�������ǰ���������飬ʹ��ArrayParser������
            Call ArrayParser(NewVarNodesIndex, VarName & " = " & VarValue)
        ElseIf InStr(VarValue, " = ") = 0 Then                                  '��������Ⱥ�û�г��ֹ���˵����������ֵ���飨a = {1, 2, ...}����ʹ��ArrayParser������
            Call ArrayParser(NewVarNodesIndex, VarName & " = " & VarValue)
        Else                                                                    '������һ���ṹ�����߳���һЩû�д���������ʹ��BracketsParser������
            Call BracketsParser(NewVarNodesIndex, VarName & " = " & VarValue)
        End If
    End If
    If VarValue Like "(* *) 0x*" Then                                       '������������������ֵ�������ڡ�(type *) 0xabc����ָ���������ֻ������ֵַ
        VarValue = Right(VarValue, Len(VarValue) - InStr(VarValue, "*) ") - 2)  '��(type *****) [0xabcde]��
    End If
    
    StringParser = Len(VarName) + 3 + Len(VarValue)                         '�����������ĳ���
End Function

'����:      �������硰x = {a, b, c, ...}��������������
'����:      ParentItem: ĸ�ڵ����
'.          OutputString: ��Ҫ�������ַ���
Private Sub ArrayParser(ParentItem As Long, OutputString As String)
    'On Error Resume Next       'todo
    Dim SplitTmp()              As String                                   '�ַ����ָ��
    Dim tmpStr                  As String                                   '�ַ���������
    Dim NewVarNodesIndex        As Long                                     '����ӵ�VarNodesԪ������
    Dim VarName                 As String                                   '����������
    Dim VarTypeName             As String                                   '����������
    Dim PipeOutput              As String                                   'gdb�ܵ����
    Dim BracketStartPos         As Long                                     '��ά������ÿ����{����λ��
    Dim BracketLevel            As Long                                     '������ƥ�������һ��ʼ��0��������{����1��������}����1
    Dim ArrayElementIndex       As Long                                     '�����ά�����ʱ��������¼����Ԫ�ص�����
    Dim NewParentVarNodesIndex  As Long                                     '�����ά�����ʱ��������¼����ӵ�VarNodesԪ������
    Dim StartQuotePos           As Long                                     '�����ַ��������ʱ�򣬲��ҵ��Ŀ�ͷ�ġ�"����λ��
    Dim EndQuotePos             As Long                                     '�����ַ��������ʱ�򣬲��ҵ��Ľ�β�ġ�"����λ��
    Dim i                       As Long
    
    If ParentItem = -1 Then                                                                     '����ǵ�һ��ı��������������Ŀ
        StringParser ParentItem, OutputString, NewVarNodesIndex
    Else                                                                                        '����Ͳ������Ŀ��ֻ��¼ĸ�ڵ�����
        NewVarNodesIndex = ParentItem
    End If
    SplitTmp = Split(OutputString, " = ")                                                       'ͨ���� = �����ָ��ַ���
    VarName = SplitTmp(0)                                                                       '��ȡ�� = ����ߵı�����
    
    '��ȡ��������
    frmMain.GdbPipe.ClearPipe                                                                   '��չܵ��������
    frmMain.GdbPipe.DosInput "ptype (" & VarName & "[0])" & vbCrLf                              '��gdb���ͻ�ȡ���ʽ�������ptype (var[0])��
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                                              '��ȡgdb���
    PipeOutput = Left(PipeOutput, InStrRev(PipeOutput, vbCrLf))                                 'ȥ����������һ��
    If Left(PipeOutput, 7) = "type = " Then                                                     '��������ڡ�type = *��
        VarTypeName = Right(PipeOutput, Len(PipeOutput) - 7)                                        '��type = [*]��
    Else
        VarTypeName = Lang_Locals_Error
    End If
    
    tmpStr = Right(OutputString, Len(OutputString) - Len(SplitTmp(0)) - 3)                      '��var = [{a, b, c, ...}]��
    tmpStr = Left(Right(tmpStr, Len(tmpStr) - 1), Len(tmpStr) - 2)                              '��{[a, b, c, ...]}��
    StartQuotePos = InStr(tmpStr, """") - 1                                                     '���ҵ�һ����"����λ��
    BracketStartPos = 1                                                                         '��ʼ����һ����{����λ��
    
    'ToDo: handle ��, <incomplete sequence \214>,��
    VarName = VarName & "["                                                                     '������������ϡ�[����Ϊ֮���������Ԫ�������׼��
    If StartQuotePos >= 0 Then                                                                  '����ҵ��ˡ�"����������һ�����ж�
        If Left(tmpStr, StartQuotePos) = String(StartQuotePos, "{") Then                            '����һ���ַ������飨var = {{...(n��{)...{"*��
            If StartQuotePos > 0 Then                                                                   '����Ƕ�ά���飬�ͼ���ʹ��ArrayParser������
                For i = 2 To Len(tmpStr)                                                                    '���ҵ�һ����{����ƥ�����һ����}��
                    If Mid(tmpStr, i, 1) = "{" Then
                        BracketLevel = BracketLevel + 1
                    ElseIf Mid(tmpStr, i, 1) = "}" Then
                        If BracketLevel <= 0 Then                                                               '���ҵ�ƥ��ġ�}������ʱi����һ��ƥ��ġ�}����λ��
                            NewParentVarNodesIndex = AddVarItem(NewVarNodesIndex, VarName & ArrayElementIndex & "]", _
                                VarName & ArrayElementIndex & "]", VarTypeName, _
                                Mid(tmpStr, BracketStartPos, i - BracketStartPos + 1))
                            Call ArrayParser(NewParentVarNodesIndex, _
                                VarName & ArrayElementIndex & "] = " & Mid(tmpStr, BracketStartPos, i - BracketStartPos + 1))
                            ArrayElementIndex = ArrayElementIndex + 1
                            BracketStartPos = i + 3                                                                 '������, ��
                            BracketLevel = -1
                        Else
                            BracketLevel = BracketLevel - 1
                        End If
                    ElseIf Mid(tmpStr, i, 1) = """" Then                                                        '������"�������ҵ���һ��ƥ��ġ�"����ȷ������������ַ����м�ȥ
                        Do
                            i = i + 1
                        Loop Until (Mid(tmpStr, i, 1) = """" And Mid(tmpStr, i - 1, 1) <> "\") Or i > Len(tmpStr)   'һֱ�����ҡ�"����ֱ���������ַ����м�
                    End If
                Next i
            Else                                                                                        '����������������Ԫ��
                StartQuotePos = 1
                For i = 2 To Len(tmpStr)                                                                    '���ҿ�ͷ�ġ�"����Ӧ����һ����"��
                    Do                                                                                          'һֱ�����ҡ�"����ֱ���������ַ����м�
                        i = i + 1
                    Loop Until (Mid(tmpStr, i, 1) = """" And Mid(tmpStr, i - 1, 1) <> "\") Or i > Len(tmpStr)   '������\"�����������ַ�������ġ�"��
                    Call AddVarItem(NewVarNodesIndex, VarName & ArrayElementIndex & "]", VarName & ArrayElementIndex & "]", _
                        VarTypeName, Mid(tmpStr, StartQuotePos + 1, i - StartQuotePos - 1))
                    ArrayElementIndex = ArrayElementIndex + 1
                    i = i + 3                                                                                   '������, ��
                    StartQuotePos = i                                                                           '��¼�µġ�"����λ��
                Next i
            End If
            Exit Sub                                                                                    '�˳����̣���Ҫִ���������ֵ�������
        End If
    End If
        
    '�ⲻ��һ���ַ�������
    If UBound(SplitTmp) = 1 Then                                                                '�Ⱥ�ֻ������һ�Σ�˵���ǲ�����������飨��ֵ���飩
        SplitTmp = Split(tmpStr, ", ")                                                              '�ԡ�, ������
        If Left(SplitTmp(0), 1) = "{" Then                                                          '����Ƕ�ά���飬�ͼ���ʹ��ArrayParser������
            For i = 2 To Len(tmpStr)                                                                    '���ҵ�һ����{����ƥ�����һ����}��
                If Mid(tmpStr, i, 1) = "{" Then
                    BracketLevel = BracketLevel + 1
                ElseIf Mid(tmpStr, i, 1) = "}" Then
                    If BracketLevel <= 0 Then                                                               '���ҵ�ƥ��ġ�}������ʱi����һ��ƥ��ġ�}����λ��
                        NewParentVarNodesIndex = AddVarItem(NewVarNodesIndex, VarName & ArrayElementIndex & "]", _
                            VarName & ArrayElementIndex & "]", VarTypeName, _
                            Mid(tmpStr, BracketStartPos, i - BracketStartPos + 1))
                        Call ArrayParser(NewParentVarNodesIndex, _
                            VarName & ArrayElementIndex & "] = " & Mid(tmpStr, BracketStartPos, i - BracketStartPos + 1))
                        ArrayElementIndex = ArrayElementIndex + 1
                        BracketStartPos = i + 3
                        BracketLevel = -1
                    Else                                                                                    '���ҵ��ġ�}�����ڲ�ģ�������1������������һ��
                        BracketLevel = BracketLevel - 1
                    End If
                End If
            Next i
        Else                                                                                        '����������������Ԫ��
            For i = 0 To UBound(SplitTmp)                                                               '�������Ԫ��
                Call AddVarItem(NewVarNodesIndex, VarName & i & "]", VarName & i & "]", VarTypeName, SplitTmp(i))
            Next i
        End If
    Else                                                                                        '�Ⱥų����˶�Σ�˵����ĳ���������
        For i = 2 To Len(tmpStr)                                                                    '���ҵ�һ����{����ƥ�����һ����}��
            If Mid(tmpStr, i, 1) = "{" Then
                BracketLevel = BracketLevel + 1
            ElseIf Mid(tmpStr, i, 1) = "}" Then
                If BracketLevel <= 0 Then                                                               '���ҵ�ƥ��ġ�}������ʱi����һ��ƥ��ġ�}����λ��
                    NewParentVarNodesIndex = AddVarItem(NewVarNodesIndex, VarName & ArrayElementIndex & "]", _
                        VarName & ArrayElementIndex & "]", VarTypeName, _
                        Mid(tmpStr, BracketStartPos, i - BracketStartPos + 1))
                    Call BracketsParser(NewParentVarNodesIndex, _
                        VarName & ArrayElementIndex & "] = " & Mid(tmpStr, BracketStartPos, i - BracketStartPos + 1))
                    ArrayElementIndex = ArrayElementIndex + 1
                    BracketStartPos = i + 3
                    BracketLevel = -1
                Else                                                                                    '���ҵ��ġ�}�����ڲ�ģ�������1������������һ��
                    BracketLevel = BracketLevel - 1
                End If
            End If
        Next i
    End If
End Sub

'����:      �������硰x = {..., ..., ...}�������
'����:      ParentItem: ĸ�ڵ����
'.          OutputString: ��Ҫ�������ַ���
'.          Start: ��ʼ������λ��
Private Sub BracketsParser(ParentItem As Long, OutputString As String)
    'On Error Resume Next       'todo
    Dim VarOutputLength         As Long                                     '�ñ���������ĳ���
    Dim SplitTmp()              As String                                   '�ַ����ָ��
    Dim tmpStr                  As String                                   '�ַ���������
    Dim NewVarNodesIndex        As Long                                     '����ӵ�VarNodesԪ������
    
    If ParentItem = -1 Then                                                                     '����ǵ�һ��ı��������������Ŀ
        VarOutputLength = StringParser(ParentItem, OutputString, NewVarNodesIndex)
        Me.picSelMargin.PaintPicture Me.imgFolded.Picture, 0, VarNodes(NewVarNodesIndex).ListViewItemIndex * ListItemHeight + 60
    Else                                                                                        '����Ͳ������Ŀ��ֻ��¼ĸ�ڵ�����
        VarOutputLength = Len(OutputString)
        NewVarNodesIndex = ParentItem
    End If
    SplitTmp = Split(Mid(OutputString, 1, VarOutputLength), " = ")
    tmpStr = Right(OutputString, Len(OutputString) - Len(SplitTmp(0)) - 4)                      '��var = {[..., ..., ...}]��
    tmpStr = Left(tmpStr, Len(tmpStr) - 1)                                                      '��[..., ..., ...]��
    Do
        tmpStr = SplitTmp(0) & "." & tmpStr                                                     '��[ParentVar.]*, *, *��
        VarOutputLength = StringParser(NewVarNodesIndex, tmpStr)
        If Len(tmpStr) > VarOutputLength + 2 Then
            tmpStr = Right(tmpStr, Len(tmpStr) - VarOutputLength - 2)                               '2 = Len(", ")
        Else
            Exit Do
        End If
    Loop
End Sub

'����:      ��ȡ���ر����б�
Public Sub GetLocals()
    'On Error Resume Next       'todo
    Dim PipeOutput              As String                                   '�ܵ������
    Dim OutputLines()           As String                                   '�����ÿһ��
    Dim VarInfoSplitTemp()      As String                                   '�������ÿһ�еķָ��
    Dim i                       As Long
    
    ReDim VarNodes(0)                                                       '��ʼ�������б�
    Me.lvLocals.Clear
    frmMain.DockingPane.Panes(8).Title = Lang_Locals_Retrieving_Caption
    
    frmMain.GdbPipe.ClearPipe                                               '��չܵ��������
    frmMain.GdbPipe.DosInput "info locals" & vbCrLf                         '��gdb���ͻ�ȡ���ر�������
    frmMain.GdbPipe.DosInput "info args" & vbCrLf                           '��gdb���ͻ�ȡ������������
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                          '��ȡgdb���
    
    OutputLines = Split(PipeOutput, vbCrLf)                                 '���зָ���
    For i = 0 To UBound(OutputLines)                                        '���н��з���
        If Trim(OutputLines(i)) <> "(gdb)" Then                                 'ȥ�����������(gdb) ��
            If Left(OutputLines(i), 6) = "(gdb) " Then                              'ȥ��һЩ�����ͷ�ġ�(gdb) ����(gdb) [*]��
                OutputLines(i) = Right(OutputLines(i), Len(OutputLines(i)) - 6)
            End If
            VarInfoSplitTemp = Split(OutputLines(i), " = ")                         '[*] = [*]
            If UBound(VarInfoSplitTemp) > 0 Then                                    '��ʱ��gdb�������No arguments.�������ﴦ���������
                If Left(VarInfoSplitTemp(1), 1) = "{" Then                              '�Ⱥŵ��ұ��ǡ�{��
                    If Left(VarInfoSplitTemp(1), 2) = "{{" Then                             '�Ⱥŵ�����ǡ�{{����˵�������ǰ���������飬ʹ��ArrayParser������
                        Call ArrayParser(-1, OutputLines(i))
                    ElseIf UBound(VarInfoSplitTemp) = 1 Then                                '�Ⱥ�ֻ������һ�Σ�˵����������ֵ���飨a = {1, 2, ...}����ʹ��ArrayParser������
                        Call ArrayParser(-1, OutputLines(i))
                    ElseIf Left(VarInfoSplitTemp(1), 2) = "{""" Then                        '�Ⱥŵ�����ǡ�{"����˵���������ַ������飬ʹ��ArrayParser������
                        Call ArrayParser(-1, OutputLines(i))
                    Else                                                                    '������һ���ṹ�����߳���һЩû�д���������ʹ��BracketsParser������
                        Call BracketsParser(-1, OutputLines(i))
                    End If
                Else                                                                    'û�г��ֵ��ںţ�Ӧ�ó����ˣ�ʹ��StringParser������
                    Call StringParser(-1, OutputLines(i))
                End If
            End If
        End If
    Next i
    
    Call RedrawNodeIcons                                                    'ˢ�½ڵ�ͼ��
    frmMain.GdbPipe.StopRecvOutput                                          'ֹͣ�ܵ����ڽ��еĹ���
    frmMain.GdbPipe.ClearPipe                                               '����β����չܵ��������
    frmMain.DockingPane.Panes(8).Title = Lang_Locals_Caption
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_Locals_Caption
    
    Me.picSelMargin.Move 0, 0, 300
    Me.lvLocals.Move 0, 0
    
    Me.lvLocals.AddColumnHeader Lang_Locals_ListViewHeader_Name, 175
    Me.lvLocals.AddColumnHeader Lang_Locals_ListViewHeader_Type, 100
    Me.lvLocals.AddColumnHeader Lang_Locals_ListViewHeader_Value
    
    ReDim VarNodes(0)                                                                               '��ʼ��VarNodes����
    
    '��ȡ�б�ͷ�ĸ߶�
    Dim tmpRect                 As RECT
    ColumnHeader = SendMessageA(Me.lvLocals.ListViewHwnd, LVM_GETHEADER, 0, 0)                      '��ȡ�б�ͷ�ľ��
    SendMessageA ColumnHeader, HDM_GETITEMRECT, ByVal 0, ByVal VarPtr(tmpRect)                      '��ȡ�б�ͷ�Ĵ�С
    ColumnHeaderHeight = (tmpRect.bottom - tmpRect.Top) * Screen.TwipsPerPixelY                     '������б�ͷ�ĸ߶�
    
    '��ȡListView��ÿ���б���ĸ߶�
    ZeroMemory tmpRect, ByVal Len(tmpRect)
    tmpRect.Left = LVIR_BOUNDS                                                                      '�����ĵ����ڷ���ϢǰtmpRect.Left������ΪLVIR_BOUNDS
    Me.lvLocals.AddItem "I need a girlfriend *(�s3�t)*"                                               '���һ���б���Լ����б���߶�
    SendMessageA Me.lvLocals.ListViewHwnd, LVM_GETITEMRECT, ByVal 0, ByVal VarPtr(tmpRect)          '��ȡ�б���Ĵ�С
    Me.lvLocals.Clear                                                                               '����б���
    ListItemHeight = (tmpRect.bottom - tmpRect.Top) * Screen.TwipsPerPixelY                         '������б���ĸ߶�
    
    '�����б�ͷ������С�Ĵ�����Ϣ���� todo
    SetPropA ColumnHeader, "PrevWndProc", SetWindowLongA(ColumnHeader, GWL_WNDPROC, AddressOf LocalsColumnHeaderLayoutProc)
    
    '����ListView�ػ�ڵ�ͼ��Ĵ�����Ϣ���� todo
    SetPropA Me.lvLocals.ListViewHwnd, "PrevWndProc", SetWindowLongA(Me.lvLocals.ListViewHwnd, GWL_WNDPROC, AddressOf LocalsListViewNodesRedrawProc)
    
    '��ͼƬ��ŵ�ListView��
    SetParent Me.picSelMargin.hWnd, Me.lvLocals.ListViewHwnd
    Me.picSelMargin.Top = ColumnHeaderHeight
    
    '����ͼƬ��Ŀ���൱�ڶ��ٸ��ո�
    SpaceCount = Me.picSelMargin.Width / Me.picSelMargin.TextWidth(" ") + 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '�ָ��б�ͷ�Ĵ�����Ϣ����
    SetWindowLongA ColumnHeader, GWL_WNDPROC, GetPropA(ColumnHeader, "PrevWndProc")
    
    '�ָ�ListView�Ĵ�����Ϣ����
    SetWindowLongA Me.lvLocals.ListViewHwnd, GWL_WNDPROC, GetPropA(Me.lvLocals.ListViewHwnd, "PrevWndProc")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.picSelMargin.Height = Me.ScaleHeight - ColumnHeaderHeight
    Me.lvLocals.Width = Me.ScaleWidth
    Me.lvLocals.Height = Me.ScaleHeight
End Sub

Private Sub lvLocals_Click(iItem As Long, iSubItem As Long, X As Long, Y As Long)
    'On Error Resume Next       'todo
    Dim i                       As Long
    
    For i = 0 To UBound(VarNodes)                                           '�����б�����ƥ���VarNodes
        If VarNodes(i).ListViewItemIndex = iItem Then
            CtlAddToolTip Me.lvLocals.ListViewHwnd, Lang_Locals_ListViewHeader_Type & ": " & VarNodes(i).TypeName & vbCrLf & _
                Lang_Locals_ListViewHeader_Value & ": " & _
                IIf(Len(VarNodes(i).Value) > 200, Left(VarNodes(i).Value, 100) & " ... " & Right(VarNodes(i).Value, 100), VarNodes(i).Value), _
                Lang_Locals_Tooltip_Title & VarNodes(i).VarName, TTI_INFO
            Exit For
        End If
    Next i
End Sub

Private Sub lvLocals_DoubleClick(iItem As Long, iSubItem As Long, X As Long, Y As Long)
    'On Error Resume Next       'todo
    Dim i                       As Long
    
    '�����б���˫��
    For i = 0 To UBound(VarNodes)                                           '�����б�����ƥ���VarNodes
        If VarNodes(i).ListViewItemIndex = iItem Then
            If VarNodes(i).Expanded Then                                            '�۵���չ���Ľڵ�
                Call FoldItem(i)
                VarNodes(i).Expanded = False
            ElseIf UBound(VarNodes(i).ChildNodes) > 0 Then                          '����ýڵ����۵����������ӽڵ㣬��չ����
                Call ExpandItem(i)
                VarNodes(i).Expanded = True
            End If
            Exit For
        End If
    Next i
End Sub

Private Sub lvLocals_KeyDown(KeyCode As Integer, Shift As Integer)
    'On Error Resume Next       'todo
    Dim ItemIndex               As Long                                     '��ǰѡ����б���
    Dim i                       As Long
    
    If KeyCode <> VK_LEFT And KeyCode <> VK_RIGHT Then                      'ֻ�������ҷ�������¼�
        Exit Sub
    End If
    
    ItemIndex = Me.lvLocals.GetSelectedItem()                               '��ȡѡ����б���
    If ItemIndex = -1 Then                                                  '��ȡʧ�ܴ���
        Exit Sub
    End If
    For i = 0 To UBound(VarNodes)                                           '�����б����Ӧ��VarNodes���
        If VarNodes(i).ListViewItemIndex = ItemIndex Then
            If KeyCode = VK_LEFT Then                                               '������������۵���ǰ�ڵ㣨ǰ���Ǹýڵ���չ����
                If VarNodes(i).Expanded Then
                    Call FoldItem(i)
                    VarNodes(i).Expanded = False
                End If
                Me.lvLocals.SetSelectedItem ItemIndex - 1                           '����һ���б����ȡ����
            ElseIf KeyCode = VK_RIGHT Then                                          '�����ҷ������չ����ǰ�ڵ㣨ǰ���Ǹýڵ�δչ�������ӽڵ㣩
                If VarNodes(i).Expanded = False And UBound(VarNodes(i).ChildNodes) > 0 Then
                    Call ExpandItem(i)
                    VarNodes(i).Expanded = True
                End If
                Me.lvLocals.SetSelectedItem ItemIndex + 1                               '����һ���б����ȡ����
            End If
            
            Exit For
        End If
    Next i
    Me.lvLocals.EnsureVisible Me.lvLocals.GetSelectedItem(), True           'ȷ����ǰѡ����б��������û�����
End Sub

Private Sub picSelMargin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim ItemIndex               As Long                                     '��갴����������Ӧ���б������
    Dim i                       As Long
    
    ItemIndex = Me.lvLocals.GetTopIndex() + Y / ListItemHeight - 1          '�������Ӧ���б������
    If ItemIndex = -1 Then
        ItemIndex = 0
    End If
    For i = 0 To UBound(VarNodes)                                           '��VarNodes�в����ĸ�Ԫ����ƥ����б������
        If VarNodes(i).ListViewItemIndex = ItemIndex Then
            If VarNodes(i).Expanded Then                                            '����ýڵ㴦��չ��״̬���Ͱ����۵�
                Call FoldItem(i)
                VarNodes(i).Expanded = False                                            '�ѵ�ǰ�ڵ���Ϊ���۵�
            Else                                                                    '����ýڵ㴦���۵�״̬���Ͱ���չ��
                Call ExpandItem(i)
                VarNodes(i).Expanded = True                                             '�ѵ�ǰ�ڵ���Ϊ��չ��
            End If
            
            Exit For
        End If
    Next i
End Sub

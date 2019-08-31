VERSION 5.00
Begin VB.Form frmLocals 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "本地"
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
'描述:      本地变量窗口，显示当前过程里的变量名称、值和类型，并允许用户对变量的值进行修改
'作者:      冰棍
'文件:      frmLocals.frm
'====================================================

Option Explicit

'定义变量信息结构
Private Type NodeInfo
    ParentNode                  As Long                                     '母节点的序号（-1代表最顶层）
    ChildNodes()                As Long                                     '所有子节点的序号（遇到-1就停止）
    VarName                     As String                                   '变量名
    Value                       As String                                   '值
    TypeName                    As String                                   '类型名
    ListViewItemIndex           As Long                                     '在ListView里面对应的列表项（-1代表没有对应的列表项）
    Expanded                    As Boolean                                  '该节点是否已展开（True = 已展开）
End Type

Dim VarNodes()                  As NodeInfo                                 '所有变量信息（最后一个元素是多余的）
Dim ColumnHeaderHeight          As Long                                     'ListView的ColumnHeader高度
Dim ListItemHeight              As Long                                     'ListView每个列表项的高度
Dim SpaceCount                  As Integer                                  '图片框的宽度相当于多少个空格
Dim ColumnHeader                As Long                                     '列表头的窗口句柄

'描述:      清空所有东西，为下一次调试做准备
Public Sub ClearEverything()
    Me.lvLocals.Clear
    ReDim VarNodes(0)
    Me.picSelMargin.Cls
End Sub

'描述:      折叠指定的节点及其子节点
'参数:      VarNodeIndex: 需要被折叠的节点序号
Private Sub FoldItem(VarNodeIndex As Long)
    'On Error Resume Next       'todo
    Dim i                       As Long
    Dim j                       As Long
    
    For i = UBound(VarNodes(VarNodeIndex).ChildNodes) - 1 To 0 Step -1      '检查下一层的子节点（倒着来）
        If VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).Expanded Then         '如果下一层子节点还有子节点，就先删掉它的子节点，再标记成已折叠
            Call FoldItem(VarNodes(VarNodeIndex).ChildNodes(i))
            VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).Expanded = False
        End If
        
        '删除掉列表项
        Me.lvLocals.DeleteItem VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).ListViewItemIndex
        VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).ListViewItemIndex = -1
        VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).Expanded = False
    Next i
    
    '被删列表项之后的列表项（即受到删除影响的列表项）向前移
    For j = VarNodeIndex To UBound(VarNodes) - 1
        If VarNodes(j).ListViewItemIndex > VarNodes(VarNodeIndex).ListViewItemIndex Then
            VarNodes(j).ListViewItemIndex = VarNodes(j).ListViewItemIndex - UBound(VarNodes(VarNodeIndex).ChildNodes)
        End If
    Next j
End Sub

'描述:      展开指定的节点
'参数:      VarNodeIndex: 需要被折叠的节点序号
Private Sub ExpandItem(VarNodeIndex As Long)
    'On Error Resume Next       'todo
    Dim Level                   As Long                                     '当前节点处于第几级
    Dim ParentIndex             As Long                                     '当前节点所对应的母节点的索引
    Dim NewItemIndex            As Long                                     '新添加的ListView列表项的索引
    Dim i                       As Long
    
    ParentIndex = VarNodeIndex
    Level = 1
    Do                                                                      '计算从当前节点到最顶层的层数
        ParentIndex = VarNodes(ParentIndex).ParentNode
        Level = Level + 1
    Loop Until ParentIndex = -1
    
    NewItemIndex = VarNodes(VarNodeIndex).ListViewItemIndex                 '令新列表项插入到当前节点的后面
    For i = VarNodeIndex + 1 To UBound(VarNodes) - 1                        '遍历VarNodes，如果其对应的列表项是在新列表项之后的就把它向后移
        If VarNodes(i).ListViewItemIndex > NewItemIndex Then
            VarNodes(i).ListViewItemIndex = VarNodes(i).ListViewItemIndex + UBound(VarNodes(VarNodeIndex).ChildNodes)
        End If
    Next i
    For i = 0 To UBound(VarNodes(VarNodeIndex).ChildNodes) - 1              '添加所有下一层子节点的列表项
        NewItemIndex = Me.lvLocals.AddItem(Space(SpaceCount * Level) & VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).VarName, NewItemIndex + 1)
        VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).ListViewItemIndex = NewItemIndex
        Me.lvLocals.SetItemText VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).TypeName, NewItemIndex, 1
        Me.lvLocals.SetItemText VarNodes(VarNodes(VarNodeIndex).ChildNodes(i)).Value, NewItemIndex, 2
    Next i
End Sub

'描述:      重绘所有的节点图标
Public Sub RedrawNodeIcons()
    On Error Resume Next
    Dim i                       As Long
    Dim TopItem                 As Long, BottomItem                 As Long

    Me.picSelMargin.Cls
    TopItem = Me.lvLocals.GetTopIndex()                                     '获取ListView中第一个可视的列表项的序号
    BottomItem = TopItem + Me.lvLocals.Height / ListItemHeight              '计算出ListView中最后一个可视的列表项的序号
    For i = 0 To UBound(VarNodes) - 1
         If VarNodes(i).ChildNodes(0) <> -1 Then                                 '如果这个变量有子节点
             If VarNodes(i).ListViewItemIndex >= TopItem And _
                VarNodes(i).ListViewItemIndex <= BottomItem Then                     '并且在可视的范围内

                 Me.picSelMargin.PaintPicture IIf(VarNodes(i).Expanded, Me.imgExpanded.Picture, Me.imgFolded.Picture), _
                     0, (VarNodes(i).ListViewItemIndex - TopItem) * ListItemHeight + 60
             End If
         End If
    Next i
End Sub

'描述:      往VarNodes和ListView里面添加项目
'参数:      ParentNode: 如果是最顶层，则为-1; 否则为母节点的序号
'.          VarName: 实际的变量名
'.          DisplayName: 显示的名称
'.          TypeName: 变量的类型
'.          Value: 变量的值
'返回值:    新添加的VarNodes元素索引
Private Function AddVarItem(ParentNode As Long, VarName As String, DisplayName As String, TypeName As String, Value As String) As Long
    'On Error Resume Next       'todo
    Dim NewItemIndex            As Long
    
    NewItemIndex = UBound(VarNodes)
    With VarNodes(NewItemIndex)                                             '设置变量信息
        .ParentNode = ParentNode
        .TypeName = TypeName
        .Value = Value
        .VarName = VarName
        ReDim .ChildNodes(0)                                                    '初始化子节点数组
        .ChildNodes(0) = -1                                                     '-1表示没有子节点
    End With
    AddVarItem = NewItemIndex
    ReDim Preserve VarNodes(NewItemIndex + 1)                               '扩充变量信息数组
    If ParentNode <> -1 Then                                                '如果新的项目不在第一层
        With VarNodes(ParentNode)
            .ChildNodes(UBound(.ChildNodes)) = NewItemIndex                         '扩充该项目的母节点的子节点数组，并把该项目的序号记录进去
            ReDim Preserve .ChildNodes(UBound(.ChildNodes) + 1)
            .ChildNodes(UBound(.ChildNodes)) = -1                                   '把新扩充的元素设置为-1，代表没有子节点
        End With
        VarNodes(NewItemIndex).ListViewItemIndex = -1                           '把对应的列表项需要设置为-1，代表没有对应的列表项
    Else                                                                    '如果新的项目在第一层，就把它添加到ListView里
        '添加列表项并设置文本
         VarNodes(NewItemIndex).ListViewItemIndex = Me.lvLocals.AddItem(Space(SpaceCount) & DisplayName)
         NewItemIndex = VarNodes(NewItemIndex).ListViewItemIndex
         Me.lvLocals.SetItemText TypeName, NewItemIndex, 1
         Me.lvLocals.SetItemText Value, NewItemIndex, 2
    End If
End Function

'描述:      分析形如“x = 0x12345 "*"”的字符串变量输出
'参数:      ParentItem: 母节点序号
'.          OutputString: 需要分析的字符串
'.          NewVarNodesIndex: 返回新添加的VarNodes元素序号
'返回值:    变量输出的长度
Private Function StringParser(ParentItem As Long, OutputString As String, Optional ByRef NewVarNodesIndex As Long) As Long
    'On Error Resume Next       'todo
    Dim VarName                 As String                                   '变量完整的名称（类似于a.b.c）
    Dim VarTypeName             As String                                   '变量类型
    Dim VarValue                As String                                   '变量的值
    Dim PipeOutput              As String                                   '管道的输出
    Dim SplitTmp()              As String                                   '字符串以“ = ”进行分割的缓存
    
    VarName = Split(OutputString, " = ")(0)                                 '分割出变量名
    
    '获取变量的值
    frmMain.GdbPipe.ClearPipe                                               '清空管道里的内容
    frmMain.GdbPipe.DosInput "p (" & VarName & ")" & vbCrLf                 '向gdb发送计算表达式命令（p (var)）
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                          '获取gdb输出
    PipeOutput = Split(PipeOutput, vbCrLf)(0)                               '只需要输出的第一行
    SplitTmp = Split(PipeOutput, " = ")                                     '（[var] = [value]）
    If UBound(SplitTmp) > 0 Then                                            '如果输出是类似于“* = *”
        VarValue = Right(PipeOutput, Len(PipeOutput) - Len(SplitTmp(0)) - 3)    '（var = [value]）
    Else                                                                    '不寻常的输出
        VarValue = Lang_Locals_Error
    End If
        
    '获取变量类型
    frmMain.GdbPipe.ClearPipe                                               '清空管道里的内容
    frmMain.GdbPipe.DosInput "ptype (" & VarName & ")" & vbCrLf             '向gdb发送获取表达式类型命令（ptype (var)）
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                          '获取gdb输出
    PipeOutput = Left(PipeOutput, InStrRev(PipeOutput, vbCrLf))             '去掉输出的最后一行
    If Left(PipeOutput, 7) = "type = " Then                                 '输出类似于“type = *”
        VarTypeName = Right(PipeOutput, Len(PipeOutput) - 7)                    '（type = [*]）
    Else
        VarTypeName = Lang_Locals_Error
    End If
    
    SplitTmp = Split(VarName, ".")                                          '以“.”分割变量名（[a].[b].[c]），然后只保留最后一个名称（a.b.[c]）
    NewVarNodesIndex = AddVarItem(ParentItem, VarName, SplitTmp(UBound(SplitTmp)), VarTypeName, VarValue)
                
    If Left(VarValue, 1) = "{" And ParentItem <> -1 Then                    '特殊情况：变量的值以“{”开头，并且非第一层变量，说明能进一步处理
        If Left(VarValue, 2) = "{{" Then                                        '变量的左边是“{{”，说明可能是包含类的数组，使用ArrayParser来处理
            Call ArrayParser(NewVarNodesIndex, VarName & " = " & VarValue)
        ElseIf InStr(VarValue, " = ") = 0 Then                                  '变量里面等号没有出现过，说明可能是数值数组（a = {1, 2, ...}），使用ArrayParser来处理
            Call ArrayParser(NewVarNodesIndex, VarName & " = " & VarValue)
        Else                                                                    '可能是一个结构，或者出现一些没有处理的情况，使用BracketsParser来处理
            Call BracketsParser(NewVarNodesIndex, VarName & " = " & VarValue)
        End If
    End If
    If VarValue Like "(* *) 0x*" Then                                       '特殊情况：如果变量的值是类似于“(type *) 0xabc”的指针输出，就只保留地址值
        VarValue = Right(VarValue, Len(VarValue) - InStr(VarValue, "*) ") - 2)  '（(type *****) [0xabcde]）
    End If
    
    StringParser = Len(VarName) + 3 + Len(VarValue)                         '计算变量输出的长度
End Function

'描述:      分析形如“x = {a, b, c, ...}”的数组变量输出
'参数:      ParentItem: 母节点序号
'.          OutputString: 需要分析的字符串
Private Sub ArrayParser(ParentItem As Long, OutputString As String)
    On Error Resume Next       'todo
    Dim SplitTmp()              As String                                   '字符串分割缓存
    Dim tmpStr                  As String                                   '字符串处理缓存
    Dim NewVarNodesIndex        As Long                                     '新添加的VarNodes元素索引
    Dim VarName                 As String                                   '变量的名称
    Dim VarTypeName             As String                                   '变量的类型
    Dim PipeOutput              As String                                   'gdb管道输出
    Dim BracketStartPos         As Long                                     '多维数组里每个“{”的位置
    Dim BracketLevel            As Long                                     '大括号匹配计数，一开始是0，遇到“{”加1，遇到“}”减1
    Dim ArrayElementIndex       As Long                                     '处理多维数组的时候，用来记录数组元素的索引
    Dim NewParentVarNodesIndex  As Long                                     '处理多维数组的时候，用来记录新添加的VarNodes元素索引
    Dim StartQuotePos           As Long                                     '处理字符串数组的时候，查找到的开头的“"”的位置
    Dim EndQuotePos             As Long                                     '处理字符串数组的时候，查找到的结尾的“"”的位置
    Dim i                       As Long
    
    If ParentItem = -1 Then                                                                     '如果是第一层的变量，就先添加项目
        StringParser ParentItem, OutputString, NewVarNodesIndex
    Else                                                                                        '否则就不添加项目，只记录母节点的序号
        NewVarNodesIndex = ParentItem
    End If
    SplitTmp = Split(OutputString, " = ")                                                       '通过“ = ”来分割字符串
    VarName = SplitTmp(0)                                                                       '获取“ = ”左边的变量名
    
    '获取变量类型
    frmMain.GdbPipe.ClearPipe                                                                   '清空管道里的内容
    frmMain.GdbPipe.DosInput "ptype (" & VarName & "[0])" & vbCrLf                              '向gdb发送获取表达式类型命令（ptype (var[0])）
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                                              '获取gdb输出
    PipeOutput = Left(PipeOutput, InStrRev(PipeOutput, vbCrLf))                                 '去掉输出的最后一行
    If Left(PipeOutput, 7) = "type = " Then                                                     '输出类似于“type = *”
        VarTypeName = Right(PipeOutput, Len(PipeOutput) - 7)                                        '（type = [*]）
    Else
        VarTypeName = Lang_Locals_Error
    End If
    
    tmpStr = Right(OutputString, Len(OutputString) - Len(SplitTmp(0)) - 3)                      '（var = [{a, b, c, ...}]）
    tmpStr = Left(Right(tmpStr, Len(tmpStr) - 1), Len(tmpStr) - 2)                              '（{[a, b, c, ...]}）
    StartQuotePos = InStr(tmpStr, """") - 1                                                     '查找第一个“"”的位置
    BracketStartPos = 1                                                                         '初始化第一个“{”的位置
    
    VarName = VarName & "["                                                                     '变量名后面加上“[”，为之后添加数组元素序号做准备
    If StartQuotePos >= 0 Then                                                                  '如果找到了“"”，再做进一步的判断
        If Left(tmpStr, StartQuotePos) = String(StartQuotePos, "{") Then                            '这是一个字符串数组（var = {{...(n个{)...{"*）
            If StartQuotePos > 0 Then                                                                   '如果是多维数组，就继续使用ArrayParser来处理
                For i = 2 To Len(tmpStr)                                                                    '查找第一个“{”所匹配的下一个“}”
                    If Mid(tmpStr, i, 1) = "{" Then
                        BracketLevel = BracketLevel + 1
                    ElseIf Mid(tmpStr, i, 1) = "}" Then
                        If BracketLevel <= 0 Then                                                               '查找到匹配的“}”。此时i是下一个匹配的“}”的位置
                            If BracketStartPos >= i + 1 Then                                                        '检查括号位置是否超出查找的位置。如果超出，就退出过程
                                Exit Sub
                            End If
                            NewParentVarNodesIndex = AddVarItem(NewVarNodesIndex, VarName & ArrayElementIndex & "]", _
                                VarName & ArrayElementIndex & "]", VarTypeName, _
                                Mid(tmpStr, BracketStartPos, i - BracketStartPos + 1))
                            Call ArrayParser(NewParentVarNodesIndex, _
                                VarName & ArrayElementIndex & "] = " & Mid(tmpStr, BracketStartPos, i - BracketStartPos + 1))
                            ArrayElementIndex = ArrayElementIndex + 1
                            BracketStartPos = i + 3                                                                 '跳过“, ”
                            BracketLevel = -1
                        Else
                            BracketLevel = BracketLevel - 1
                        End If
                    ElseIf Mid(tmpStr, i, 1) = """" Then                                                        '遇到“"”，查找下一个匹配的“"”，确保不会分析到字符串中间去
                        Do
                            i = i + 1
                        Loop Until (Mid(tmpStr, i, 1) = """" And Mid(tmpStr, i - 1, 1) <> "\") Or i > Len(tmpStr)   '一直向后查找“"”，直到不处于字符串中间
                    ElseIf Mid(tmpStr, i, 1) = "<" Then                                                         '遇到“”，查找下一个匹配的“”，确保不会分析到“<>”中间去（处理新版gdb的“<incomplete sequence \*>”输出）
                        Do
                            i = i + 1
                        Loop Until Mid(tmpStr, i, 1) = ">" Or i > Len(tmpStr)                                       '一直向后查找“>”，直到不处于“<>”中间
                        i = i + 2                                                                                   '跳过“, ”
                    End If
                Next i
            Else                                                                                        '否则就依次添加所有元素
                StartQuotePos = 0
                For i = 1 To Len(tmpStr)                                                                    '查找开头的“"”对应的下一个“"”
                    If Mid(tmpStr, i, 1) = """" Then                                                            '如果是以“"”开头的元素
                        Do                                                                                          '一直向后查找“"”，直到不处于字符串中间
                            i = i + 1
                        Loop Until (Mid(tmpStr, i, 1) = """" And Mid(tmpStr, i - 1, 1) <> "\") Or i > Len(tmpStr)   '跳过“\"”，这是在字符串里面的“"”
                    ElseIf Mid(tmpStr, i, 1) = "<" Then                                                         '如果是以“<”开头的元素（处理新版gdb的“<incomplete sequence \*>”输出）
                        Do                                                                                          '一直向后查找“>”，直到不处于“<>”中间
                            i = i + 1
                        Loop Until Mid(tmpStr, i, 1) = ">" Or i > Len(tmpStr)
                    Else                                                                                        '其它东西？应该不会出现这种情况吧
                        i = Len(tmpStr)                                                                             '如果真的出现这种情况... 直接跳到结尾吧
                    End If
                    Call AddVarItem(NewVarNodesIndex, VarName & ArrayElementIndex & "]", VarName & ArrayElementIndex & "]", _
                        VarTypeName, Mid(tmpStr, StartQuotePos + 1, i - StartQuotePos))
                    ArrayElementIndex = ArrayElementIndex + 1
                    i = i + 2                                                                                   '跳过“, ”
                    StartQuotePos = i                                                                           '记录新的“"”的位置
                Next i
            End If
            Exit Sub                                                                                    '退出过程，不要执行下面的数值数组分析
        End If
    End If
        
    '这不是一个字符串数组
    If UBound(SplitTmp) = 1 Then                                                                '等号只出现了一次，说明是不含有类的数组（数值数组）
        SplitTmp = Split(tmpStr, ", ")                                                              '以“, ”隔开
        If Left(SplitTmp(0), 1) = "{" Then                                                          '如果是多维数组，就继续使用ArrayParser来处理
            For i = 2 To Len(tmpStr)                                                                    '查找第一个“{”所匹配的下一个“}”
                If Mid(tmpStr, i, 1) = "{" Then
                    BracketLevel = BracketLevel + 1
                ElseIf Mid(tmpStr, i, 1) = "}" Then
                    If BracketLevel <= 0 Then                                                               '查找到匹配的“}”。此时i是下一个匹配的“}”的位置
                        NewParentVarNodesIndex = AddVarItem(NewVarNodesIndex, VarName & ArrayElementIndex & "]", _
                            VarName & ArrayElementIndex & "]", VarTypeName, _
                            Mid(tmpStr, BracketStartPos, i - BracketStartPos + 1))
                        Call ArrayParser(NewParentVarNodesIndex, _
                            VarName & ArrayElementIndex & "] = " & Mid(tmpStr, BracketStartPos, i - BracketStartPos + 1))
                        ArrayElementIndex = ArrayElementIndex + 1
                        BracketStartPos = i + 3
                        BracketLevel = -1
                    Else                                                                                    '查找到的“}”是内层的，层数减1，继续查找下一个
                        BracketLevel = BracketLevel - 1
                    End If
                End If
            Next i
        Else                                                                                        '否则就依次添加所有元素
            For i = 0 To UBound(SplitTmp)                                                               '添加所有元素
                Call AddVarItem(NewVarNodesIndex, VarName & i & "]", VarName & i & "]", VarTypeName, SplitTmp(i))
            Next i
        End If
    Else                                                                                        '等号出现了多次，说明是某个类的数组
        For i = 2 To Len(tmpStr)                                                                    '查找第一个“{”所匹配的下一个“}”
            If Mid(tmpStr, i, 1) = "{" Then
                BracketLevel = BracketLevel + 1
            ElseIf Mid(tmpStr, i, 1) = "}" Then
                If BracketLevel <= 0 Then                                                               '查找到匹配的“}”。此时i是下一个匹配的“}”的位置
                    NewParentVarNodesIndex = AddVarItem(NewVarNodesIndex, VarName & ArrayElementIndex & "]", _
                        VarName & ArrayElementIndex & "]", VarTypeName, _
                        Mid(tmpStr, BracketStartPos, i - BracketStartPos + 1))
                    Call BracketsParser(NewParentVarNodesIndex, _
                        VarName & ArrayElementIndex & "] = " & Mid(tmpStr, BracketStartPos, i - BracketStartPos + 1))
                    ArrayElementIndex = ArrayElementIndex + 1
                    BracketStartPos = i + 3
                    BracketLevel = -1
                Else                                                                                    '查找到的“}”是内层的，层数减1，继续查找下一个
                    BracketLevel = BracketLevel - 1
                End If
            End If
        Next i
    End If
End Sub

'描述:      分析形如“x = {..., ..., ...}”的输出
'参数:      ParentItem: 母节点序号
'.          OutputString: 需要分析的字符串
'.          Start: 开始分析的位置
Private Sub BracketsParser(ParentItem As Long, OutputString As String)
    'On Error Resume Next       'todo
    Dim VarOutputLength         As Long                                     '该变量的输出的长度
    Dim SplitTmp()              As String                                   '字符串分割缓存
    Dim tmpStr                  As String                                   '字符串处理缓存
    Dim NewVarNodesIndex        As Long                                     '新添加的VarNodes元素索引
    
    If ParentItem = -1 Then                                                                     '如果是第一层的变量，就先添加项目
        VarOutputLength = StringParser(ParentItem, OutputString, NewVarNodesIndex)
        Me.picSelMargin.PaintPicture Me.imgFolded.Picture, 0, VarNodes(NewVarNodesIndex).ListViewItemIndex * ListItemHeight + 60
    Else                                                                                        '否则就不添加项目，只记录母节点的序号
        VarOutputLength = Len(OutputString)
        NewVarNodesIndex = ParentItem
    End If
    SplitTmp = Split(Mid(OutputString, 1, VarOutputLength), " = ")
    tmpStr = Right(OutputString, Len(OutputString) - Len(SplitTmp(0)) - 4)                      '（var = {[..., ..., ...}]）
    tmpStr = Left(tmpStr, Len(tmpStr) - 1)                                                      '（[..., ..., ...]）
    Do
        tmpStr = SplitTmp(0) & "." & tmpStr                                                     '（[ParentVar.]*, *, *）
        VarOutputLength = StringParser(NewVarNodesIndex, tmpStr)
        If Len(tmpStr) > VarOutputLength + 2 Then
            tmpStr = Right(tmpStr, Len(tmpStr) - VarOutputLength - 2)                               '2 = Len(", ")
        Else
            Exit Do
        End If
    Loop
End Sub

'描述:      获取本地变量列表
Public Sub GetLocals()
    'On Error Resume Next       'todo
    Dim PipeOutput              As String                                   '管道的输出
    Dim OutputLines()           As String                                   '输出的每一行
    Dim VarInfoSplitTemp()      As String                                   '对输出的每一行的分割缓存
    Dim i                       As Long
    
    ReDim VarNodes(0)                                                       '初始化变量列表
    Me.lvLocals.Clear
    frmMain.DockingPane.Panes(8).Title = Lang_Locals_Retrieving_Caption
    
    frmMain.GdbPipe.ClearPipe                                               '清空管道里的内容
    frmMain.GdbPipe.DosInput "info locals" & vbCrLf                         '向gdb发送获取本地变量命令
    frmMain.GdbPipe.DosInput "info args" & vbCrLf                           '向gdb发送获取参数变量命令
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                          '获取gdb输出
    
    OutputLines = Split(PipeOutput, vbCrLf)                                 '逐行分割开输出
    For i = 0 To UBound(OutputLines)                                        '逐行进行分析
        If Trim(OutputLines(i)) <> "(gdb)" Then                                 '去掉无用输出“(gdb) ”
            If Left(OutputLines(i), 6) = "(gdb) " Then                              '去掉一些输出开头的“(gdb) ”（(gdb) [*]）
                OutputLines(i) = Right(OutputLines(i), Len(OutputLines(i)) - 6)
            End If
            VarInfoSplitTemp = Split(OutputLines(i), " = ")                         '[*] = [*]
            If UBound(VarInfoSplitTemp) > 0 Then                                    '有时候gdb会输出“No arguments.”，这里处理这种情况
                If Left(VarInfoSplitTemp(1), 1) = "{" Then                              '等号的右边是“{”
                    If Left(VarInfoSplitTemp(1), 2) = "{{" Then                             '等号的左边是“{{”，说明可能是包含类的数组，使用ArrayParser来处理
                        Call ArrayParser(-1, OutputLines(i))
                    ElseIf UBound(VarInfoSplitTemp) = 1 Then                                '等号只出现了一次，说明可能是数值数组（a = {1, 2, ...}），使用ArrayParser来处理
                        Call ArrayParser(-1, OutputLines(i))
                    ElseIf Left(VarInfoSplitTemp(1), 2) = "{""" Then                        '等号的左边是“{"”，说明可能是字符串数组，使用ArrayParser来处理
                        Call ArrayParser(-1, OutputLines(i))
                    Else                                                                    '可能是一个结构，或者出现一些没有处理的情况，使用BracketsParser来处理
                        Call BracketsParser(-1, OutputLines(i))
                    End If
                Else                                                                    '没有出现等于号，应该出错了，使用StringParser来处理
                    Call StringParser(-1, OutputLines(i))
                End If
            End If
        End If
    Next i
    
    Call RedrawNodeIcons                                                    '刷新节点图标
    frmMain.GdbPipe.StopRecvOutput                                          '停止管道正在进行的工作
    frmMain.GdbPipe.ClearPipe                                               '捡手尾：清空管道里的内容
    frmMain.DockingPane.Panes(8).Title = Lang_Locals_Caption
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_Locals_Caption
    
    Me.picSelMargin.Move 0, 0, 300
    Me.lvLocals.Move 0, 0
    
    Me.lvLocals.AddColumnHeader Lang_Locals_ListViewHeader_Name, 175
    Me.lvLocals.AddColumnHeader Lang_Locals_ListViewHeader_Type, 100
    Me.lvLocals.AddColumnHeader Lang_Locals_ListViewHeader_Value
    
    ReDim VarNodes(0)                                                                               '初始化VarNodes数组
    
    '获取列表头的高度
    Dim tmpRect                 As RECT
    ColumnHeader = SendMessageA(Me.lvLocals.ListViewHwnd, LVM_GETHEADER, 0, 0)                      '获取列表头的句柄
    SendMessageA ColumnHeader, HDM_GETITEMRECT, ByVal 0, ByVal VarPtr(tmpRect)                      '获取列表头的大小
    ColumnHeaderHeight = (tmpRect.bottom - tmpRect.Top) * Screen.TwipsPerPixelY                     '计算出列表头的高度
    
    '获取ListView中每个列表项的高度
    ZeroMemory tmpRect, ByVal Len(tmpRect)
    tmpRect.Left = LVIR_BOUNDS                                                                      '根据文档，在发消息前tmpRect.Left须设置为LVIR_BOUNDS
    Me.lvLocals.AddItem "I need a girlfriend *(╯3╰)*"                                               '添加一个列表项，以计算列表项高度
    SendMessageA Me.lvLocals.ListViewHwnd, LVM_GETITEMRECT, ByVal 0, ByVal VarPtr(tmpRect)          '获取列表项的大小
    Me.lvLocals.Clear                                                                               '清空列表项
    ListItemHeight = (tmpRect.bottom - tmpRect.Top) * Screen.TwipsPerPixelY                         '计算出列表项的高度
    
    '设置列表头调整大小的窗口消息处理 todo
    SetPropA ColumnHeader, "PrevWndProc", SetWindowLongA(ColumnHeader, GWL_WNDPROC, AddressOf LocalsColumnHeaderLayoutProc)
    
    '设置ListView重绘节点图标的窗口消息处理 todo
    SetPropA Me.lvLocals.ListViewHwnd, "PrevWndProc", SetWindowLongA(Me.lvLocals.ListViewHwnd, GWL_WNDPROC, AddressOf LocalsListViewNodesRedrawProc)
    
    '把图片框放到ListView里
    SetParent Me.picSelMargin.hWnd, Me.lvLocals.ListViewHwnd
    Me.picSelMargin.Top = ColumnHeaderHeight
    
    '计算图片框的宽度相当于多少个空格
    SpaceCount = Me.picSelMargin.Width / Me.picSelMargin.TextWidth(" ") + 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '恢复列表头的窗口消息处理
    SetWindowLongA ColumnHeader, GWL_WNDPROC, GetPropA(ColumnHeader, "PrevWndProc")
    
    '恢复ListView的窗口消息处理
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
    
    For i = 0 To UBound(VarNodes)                                           '查找列表项所匹配的VarNodes
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
    
    '处理列表项双击
    For i = 0 To UBound(VarNodes)                                           '查找列表项所匹配的VarNodes
        If VarNodes(i).ListViewItemIndex = iItem Then
            If VarNodes(i).Expanded Then                                            '折叠已展开的节点
                Call FoldItem(i)
                VarNodes(i).Expanded = False
            ElseIf UBound(VarNodes(i).ChildNodes) > 0 Then                          '如果该节点已折叠，而且有子节点，就展开它
                Call ExpandItem(i)
                VarNodes(i).Expanded = True
            End If
            Exit For
        End If
    Next i
End Sub

Private Sub lvLocals_KeyDown(KeyCode As Integer, Shift As Integer)
    'On Error Resume Next       'todo
    Dim ItemIndex               As Long                                     '当前选择的列表项
    Dim i                       As Long
    
    If KeyCode <> VK_LEFT And KeyCode <> VK_RIGHT Then                      '只处理左、右方向键的事件
        Exit Sub
    End If
    
    ItemIndex = Me.lvLocals.GetSelectedItem()                               '获取选择的列表项
    If ItemIndex = -1 Then                                                  '获取失败处理
        Exit Sub
    End If
    For i = 0 To UBound(VarNodes)                                           '查找列表项对应的VarNodes序号
        If VarNodes(i).ListViewItemIndex = ItemIndex Then
            If KeyCode = VK_LEFT Then                                               '按下左方向键，折叠当前节点（前提是该节点已展开）
                If VarNodes(i).Expanded Then
                    Call FoldItem(i)
                    VarNodes(i).Expanded = False
                End If
                Me.lvLocals.SetSelectedItem ItemIndex - 1                           '让上一个列表项获取焦点
            ElseIf KeyCode = VK_RIGHT Then                                          '按下右方向键，展开当前节点（前提是该节点未展开且有子节点）
                If VarNodes(i).Expanded = False And UBound(VarNodes(i).ChildNodes) > 0 Then
                    Call ExpandItem(i)
                    VarNodes(i).Expanded = True
                End If
                Me.lvLocals.SetSelectedItem ItemIndex + 1                               '让下一个列表项获取焦点
            End If
            
            Exit For
        End If
    Next i
    Me.lvLocals.EnsureVisible Me.lvLocals.GetSelectedItem(), True           '确保当前选择的列表项能让用户看到
End Sub

Private Sub picSelMargin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim ItemIndex               As Long                                     '鼠标按下坐标所对应的列表项序号
    Dim i                       As Long
    
    ItemIndex = Me.lvLocals.GetTopIndex() + Y / ListItemHeight - 1          '计算出对应的列表项序号
    If ItemIndex = -1 Then
        ItemIndex = 0
    End If
    For i = 0 To UBound(VarNodes)                                           '在VarNodes中查找哪个元素有匹配的列表项序号
        If VarNodes(i).ListViewItemIndex = ItemIndex Then
            If VarNodes(i).Expanded Then                                            '如果该节点处于展开状态，就把他折叠
                Call FoldItem(i)
                VarNodes(i).Expanded = False                                            '把当前节点标记为已折叠
            Else                                                                    '如果该节点处于折叠状态，就把他展开
                Call ExpandItem(i)
                VarNodes(i).Expanded = True                                             '把当前节点标记为已展开
            End If
            
            Exit For
        End If
    Next i
End Sub

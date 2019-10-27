VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmCodeWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "代码窗口"
   ClientHeight    =   5175
   ClientLeft      =   3540
   ClientTop       =   3060
   ClientWidth     =   8865
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmCodeWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   8865
   Begin XtremeSyntaxEdit.SyntaxEdit SyntaxEdit 
      Height          =   1935
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
      _Version        =   983043
      _ExtentX        =   5318
      _ExtentY        =   3413
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   0   'False
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin VB.Timer tmrUpdateBreakpoints 
      Interval        =   50
      Left            =   6960
      Top             =   4560
   End
   Begin VB.PictureBox picSelMargin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00333333&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   240
      ScaleHeight     =   1935
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   1320
      Width           =   255
   End
   Begin DragControlsIDE.DarkComboBox comObject 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      Items0          =   ""
      ITEM_COUNT      =   0
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   7560
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      Thickness       =   4
      MinWidth        =   150
      MinHeight       =   100
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "代码窗口"
      MaxButtonVisible=   0   'False
      MinButtonVisible=   0   'False
      BindCaption     =   -1  'True
      Picture         =   "frmCodeWindow.frx":1BCC2
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorderSizer 
      Left            =   8280
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      Thickness       =   3
      FocusedColor    =   3157293
      NotFocusedColor =   3157293
      MinWidth        =   150
      MinHeight       =   100
   End
   Begin DragControlsIDE.DarkComboBox comEvent 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   660
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      Items0          =   ""
      ITEM_COUNT      =   0
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgDisabledBreakpoint 
      Height          =   240
      Left            =   6120
      Picture         =   "frmCodeWindow.frx":1C914
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgBreakpoint 
      Height          =   240
      Left            =   5760
      Picture         =   "frmCodeWindow.frx":1CC9E
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCurrentLine 
      Height          =   240
      Left            =   5400
      Picture         =   "frmCodeWindow.frx":1D028
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmCodeWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'描述:      代码窗口，负责大部分的代码相关的工作
'作者:      冰棍
'文件:      frmCodeWindow.frm
'====================================================

Option Explicit

Public WindowObj    As Object                                                       '窗口自身
Public FileIndex    As Long                                                         '在CurrentProject.Files对应的文件序号
Public RowHeight    As Integer                                                      '代码行的高度（用于计算断点绘图位置）
Public BreakLine    As Long                                                         '在调试期间中断的行（-1代表没有）

'描述:      重新通过代码框的字体计算每行代码的高度
Public Sub ReCalcRowHeight()
    Set Me.picSelMargin.Font = Me.SyntaxEdit.Font
    RowHeight = Me.picSelMargin.TextHeight("#")
End Sub

'描述:      重绘所有的断点
Public Sub RedrawBreakpoints()
    Dim lnStart     As Long, lnEnd      As Long, ln         As Long                 '可视的第一行; 可视的最后一行; 断点对应的行
    Dim i           As Long
    
    Me.picSelMargin.Cls                                                             '清空画布
    lnStart = Me.SyntaxEdit.TopRow                                                  '获取当前可视的第一行
    lnEnd = lnStart + Me.SyntaxEdit.Height / RowHeight                              '根据文本框的高度和每行的高度算出文本框能装下多少行，从而得到可视的最后一行
    If lnEnd > Me.SyntaxEdit.RowsCount Then                                         '如果可视的最后一行超过了文本框的总行数就取总行数
        lnEnd = Me.SyntaxEdit.RowsCount
    End If
    For i = 0 To UBound(CurrentProject.Files(FileIndex).Breakpoints)                '遍历当前文件的断点，如果是在可视的行数范围内的就画出来
        ln = CurrentProject.Files(FileIndex).Breakpoints(i).CodeLn
        If ln >= lnStart And ln <= lnEnd Then
            Me.picSelMargin.PaintPicture Me.imgBreakpoint.Picture, 0, RowHeight * (ln - lnStart), 240, 240
        End If
    Next i
    
    If BreakLine >= lnStart And BreakLine <= lnEnd Then
        Me.picSelMargin.PaintPicture Me.imgCurrentLine.Picture, 0, RowHeight * (BreakLine - lnStart), 240, 240
    End If
End Sub

Private Sub DarkTitleBar_GotFocus()
    On Error Resume Next
    
    Me.SyntaxEdit.SetFocus
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_CodeWindow_Caption
    Me.DarkTitleBar.MaxButtonVisible = True
    Me.DarkTitleBar.MinButtonVisible = True
    
    '设置代码框属性
    Me.DarkTitleBar.Top = Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelY
    Me.picSelMargin.Move Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelX, Me.SyntaxEdit.Top, 300, Me.SyntaxEdit.Height
    Me.SyntaxEdit.Move Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelX + Me.picSelMargin.Width, _
        Me.DarkTitleBar.Height + Me.comObject.Height + 240 + Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelY
    Me.SyntaxEdit.PaintManager.BackColor = RGB(28, 28, 28)
    Me.SyntaxEdit.PaintManager.LineNumberBackColor = RGB(28, 28, 28)
    Me.SyntaxEdit.PaintManager.LineNumberTextColor = RGB(86, 156, 214)
    Me.SyntaxEdit.DataManager.FileExt = ".cpp"
    Me.SyntaxEdit.ConfigFile = App.Path & "\SyntaxEdit.ini"
    Call ReCalcRowHeight                                                                                                '重新计算代码行高度
    
    '设置窗口子类化，处理最大化问题及处理任务栏右键关闭
    Dim lpObj               As Long                                                                                     '指向窗口自身的物件指针
    Set WindowObj = Me
    lpObj = ObjPtr(WindowObj)                                                                                           '获取指向窗口自身的物件指针
    SetPropA Me.hwnd, "WindowObj", lpObj                                                                                '记录窗口的物件地址，供子类化卸载窗体用
    'SetPropA Me.hWnd, "PrevWndProc", SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf MainWindowMaximizeCloseFixProc)    '[ToDo]

    '设置代码框的子类化，使其重绘的时候能够重绘断点
    Dim RealSyntaxEdit      As Long                                                                                     '代码框真实的hWnd
    
    RealSyntaxEdit = FindWindowExA(Me.SyntaxEdit.hwnd, 0, "CodejockSyntaxEditor", vbNullString)                         '代码框其实只是一个壳，里面的那个窗口才是真正的代码框窗口
    SetPropA RealSyntaxEdit, "FileIndex", FileIndex
    'SetPropA RealSyntaxEdit, "PrevWndProc", SetWindowLongA(RealSyntaxEdit, GWL_WNDPROC, AddressOf EditBreakpointsRedrawProc)    '[ToDo]
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsExiting Then
        '恢复窗口子类化
        SetWindowLongA Me.hwnd, GWL_WNDPROC, GetPropA(Me.hwnd, "PrevWndProc")
        SetWindowLongA Me.SyntaxEdit.hwnd, GWL_WNDPROC, GetPropA(Me.SyntaxEdit.hwnd, "PrevWndProc")
    Else
        Cancel = 1
        Me.Hide
    End If
End Sub

Public Sub Form_Resize()
    On Error Resume Next
    
    '根据标题栏是否显示来调整控件位置
    If Me.DarkTitleBar.Visible = True Then
        Me.comObject.Top = Me.DarkTitleBar.Height + 165
        Me.comEvent.Top = Me.comObject.Top
        Me.SyntaxEdit.Top = Me.comEvent.Top + Me.comEvent.Height + 240
    Else
        Me.comObject.Top = 120
        Me.comEvent.Top = 120
        Me.SyntaxEdit.Top = 120 + Me.comObject.Height + 120
    End If
    Me.picSelMargin.Top = Me.SyntaxEdit.Top
    
    '设置代码框大小
    Me.SyntaxEdit.Width = Me.ScaleWidth - Me.SyntaxEdit.Left - Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelX
    Me.SyntaxEdit.Height = Me.ScaleHeight - Me.SyntaxEdit.Top - Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelY
    Me.picSelMargin.Height = Me.SyntaxEdit.Height
    
    '设置组合框大小和位置
    Me.comObject.Width = (Me.ScaleWidth - 480) / 2
    Me.comEvent.Width = Me.comObject.Width
    Me.comEvent.Left = Me.comObject.Width + 360
End Sub

Private Sub picSelMargin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim CurrRow         As Long, CurrCol        As Long                                 '鼠标坐标对应的代码行、列。其中列是没用的，但是这个垃圾控件愣是要我传这个参数...
    Dim BreakpointCount As Long                                                         'UBound(.Breakpoints)，实际断点数量 - 1
    Dim i               As Long
    
    Me.SyntaxEdit.RowColCodeFromPoint X, Y / Screen.TwipsPerPixelY, CurrRow, CurrCol    '获取鼠标坐标对应的行
    Me.SyntaxEdit.SetFocus
    CurrentProject.Changed = True                                                       '更改断点也视为更改了文件
    
    With CurrentProject.Files(FileIndex)
        BreakpointCount = UBound(.Breakpoints)
        For i = 0 To BreakpointCount                                                    '查找对应的断点
            If .Breakpoints(i).CodeLn = CurrRow Then                                        '如果能找到对应的断点就删掉
                Dim j               As Long
                
                frmBreakpoints.lvBreakpoints.DeleteItem .Breakpoints(i).ListViewIndex           '从ListView移除对应的列表项
                For j = 0 To BreakpointCount                                                    '查找所有该列表项后面对应的断点，并把它们所对应的列表项序号 - 1
                    If .Breakpoints(j).ListViewIndex > .Breakpoints(i).ListViewIndex Then
                        .Breakpoints(j).ListViewIndex = .Breakpoints(j).ListViewIndex - 1
                    End If
                Next j
                
                If i < BreakpointCount Then                                                     '如果后面还有别的断点信息就把它们向前移
                    CopyMemory .Breakpoints(i), .Breakpoints(i + 1), LenB(.Breakpoints(0)) * (BreakpointCount - i)
                End If
                ReDim Preserve .Breakpoints(BreakpointCount - 1)                                '缩小断点数组
                Call RedrawBreakpoints                                                          '重绘所有断点
                Exit Sub
            End If
        Next i
        
        '如果不能找到对应的断点就添加
        ReDim Preserve .Breakpoints(BreakpointCount + 1)                                '扩大断点数组
        .Breakpoints(BreakpointCount).CodeLn = CurrRow                                  '设置断点对应的行数和激活状态
        .Breakpoints(BreakpointCount).Enabled = True
        .Breakpoints(BreakpointCount).ListViewIndex = frmBreakpoints.lvBreakpoints.AddItem(GetFileName(.FilePath))
        frmBreakpoints.lvBreakpoints.SetItemText CStr(CurrRow), .Breakpoints(BreakpointCount).ListViewIndex, 1
        frmBreakpoints.lvBreakpoints.SetItemChecked .Breakpoints(BreakpointCount).ListViewIndex, True
        Call RedrawBreakpoints                                                          '重绘所有断点
    End With
End Sub

Private Sub picSelMargin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim CurrRow         As Long, CurrCol        As Long                                 '鼠标坐标对应的代码行、列。其中列是没用的，但是这个垃圾控件愣是要我传这个参数...
    Dim i               As Long
    
    Me.SyntaxEdit.RowColCodeFromPoint X, Y / Screen.TwipsPerPixelY, CurrRow, CurrCol    '获取鼠标坐标对应的行
    With CurrentProject.Files(FileIndex)
        For i = 0 To UBound(.Breakpoints)                                                   '尝试查找该行有没有对应的断点
            If .Breakpoints(i).CodeLn = CurrRow Then                                            '找到匹配的断点就显示断点信息
                Me.picSelMargin.ToolTipText = Lang_Breakpoints_Info_1 & .Breakpoints(i).CodeLn & Lang_Breakpoints_Info_2 & _
                    IIf(.Breakpoints(i).Enabled, Lang_Breakpoints_Info_3, Lang_Breakpoints_Info_4)
                Exit Sub
            End If
        Next i
    End With
    Me.picSelMargin.ToolTipText = ""                                                    '找不到就啥信息都不显示
End Sub

Private Sub SyntaxEdit_TextChanged(ByVal nRowFrom As Long, ByVal nRowTo As Long, ByVal nActions As Long)
    CurrentProject.Files(FileIndex).Changed = True                                  '代码框的内容一旦更改，就把文件视为更改了
    '------------------------------------------------------
    
    Dim nLinesChanged   As Long                                                     '变化的行数
    Dim SelStartRow     As Long                                                     '选择的文本的起始行
    Dim SelEndRow       As Long                                                     '选择的文本的结束行
    Dim i               As Long
    Dim j               As Long
    
    If nRowTo <> nRowFrom Then                                                      '如果行数发生了变化
        nLinesChanged = nRowTo - nRowFrom                                               '计算行数的变化
        Select Case nActions                                                            '对一些操作进行处理
            Case 6                                                                          '删除操作（退格键，删除键，剪切等）
                nLinesChanged = -nLinesChanged
            
            Case 775, 518, 261                                                              '撤销、重复
                nLinesChanged = 0
                
        End Select
    End If
    If nLinesChanged = 0 Then                                                       '如果行数发生了变化才检查断点有没有受到影响
        Exit Sub
    End If
    SelStartRow = Me.SyntaxEdit.Selection.Start.Row
    SelEndRow = Me.SyntaxEdit.Selection.End.Row
    
    With CurrentProject.Files(FileIndex)
        For i = UBound(.Breakpoints) - 1 To 0 Step -1                                       '遍历断点列表，删除涉及的断点，并调整其它断点的位置
            If nLinesChanged < 0 And _
               ((SelEndRow <= .Breakpoints(i).CodeLn And .Breakpoints(i).CodeLn <= SelStartRow And SelEndRow < SelStartRow) Or _
               (SelStartRow <= .Breakpoints(i).CodeLn And .Breakpoints(i).CodeLn <= SelEndRow And SelStartRow <= SelEndRow)) Then
                '断点位于被删除的行中间（SelEndRow 和 SelStartRow 可以互换位置，因为用户更改的方向可以不一样）
                ' ...
                ' SelEndRow   -----  ┓
                ' ...                ┃
                '  .CodeLn    -----  ┃ 这中间的断点将被删掉
                ' ...                ┃
                ' SelStartRow -----  ┛
                ' ...
                '=====================
                '删除断点。这里的代码类似于picSelMargin_MouseDown里删除断点的代码
                frmBreakpoints.lvBreakpoints.DeleteItem .Breakpoints(i).ListViewIndex       '从ListView移除对应的列表项
                For j = 0 To UBound(.Breakpoints)                                           '查找所有该列表项后面对应的断点，并把他们所对应的列表项序号 - 1
                    If .Breakpoints(j).ListViewIndex > .Breakpoints(i).ListViewIndex Then
                        .Breakpoints(j).ListViewIndex = .Breakpoints(j).ListViewIndex - 1
                    End If
                Next j
                
                If i < UBound(.Breakpoints) Then                                            '如果后面还有别的断点信息就把它们向前移
                    CopyMemory .Breakpoints(i), .Breakpoints(i + 1), LenB(.Breakpoints(0)) * (UBound(.Breakpoints) - i)
                End If
                ReDim Preserve .Breakpoints(UBound(.Breakpoints) - 1)                       '缩小断点数组
            ElseIf .Breakpoints(i).CodeLn > nRowFrom Then
                '断点位于发生更改的行后面
                ' ...
                ' nRowFrom -----
                ' ...               ┓
                ' .CodeLn -----     ┃ 在nRowFrom下面的断点所对应的行号将被修改
                ' ...               ┃
                '=====================
                .Breakpoints(i).CodeLn = .Breakpoints(i).CodeLn + nLinesChanged
                frmBreakpoints.lvBreakpoints.SetItemText CStr(.Breakpoints(i).CodeLn), .Breakpoints(i).ListViewIndex, 1
            End If
        Next i
    End With
    
    Call RedrawBreakpoints                                                          '重绘所有断点
    bpRedrawFileIndex = -1                                                          '让计时器别重绘了
End Sub

Private Sub tmrUpdateBreakpoints_Timer()
    If bpRedrawFileIndex = FileIndex Then
        Call RedrawBreakpoints
        bpRedrawFileIndex = -1
    End If
End Sub

VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "image.ocx"
Begin VB.Form frmErrorList 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "错误列表"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTypeSelection 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00373333&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   7035
      TabIndex        =   2
      Top             =   0
      Width           =   7035
      Begin VB.Timer tmrRestoreColor 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   3960
         Top             =   120
      End
      Begin VB.PictureBox picSwitchInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00373333&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   2640
         ScaleHeight     =   405
         ScaleWidth      =   1215
         TabIndex        =   7
         Top             =   60
         Width           =   1215
         Begin ImageX.aicAlphaImage imgInfo 
            Height          =   240
            Left            =   120
            Top             =   120
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            Image           =   "frmErrorList.frx":0000
            Enabled         =   0   'False
            Props           =   13
         End
         Begin VB.Label labInfoCount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 消息"
            ForeColor       =   &H00F0F0F0&
            Height          =   195
            Left            =   480
            TabIndex        =   8
            Top             =   120
            Width           =   495
         End
         Begin VB.Shape shpBorderInfo 
            BackColor       =   &H00FF9933&
            BorderColor     =   &H00FF9933&
            FillColor       =   &H00FF9933&
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox picSwitchWarnings 
         Appearance      =   0  'Flat
         BackColor       =   &H00373333&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   1320
         ScaleHeight     =   405
         ScaleWidth      =   1215
         TabIndex        =   5
         Top             =   60
         Width           =   1215
         Begin ImageX.aicAlphaImage imgWarning 
            Height          =   240
            Left            =   120
            Top             =   120
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            Image           =   "frmErrorList.frx":02F1
            Enabled         =   0   'False
            Props           =   13
         End
         Begin VB.Label labWarningCount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 警告"
            ForeColor       =   &H00F0F0F0&
            Height          =   195
            Left            =   480
            TabIndex        =   6
            Top             =   120
            Width           =   495
         End
         Begin VB.Shape shpBorderWarnings 
            BackColor       =   &H00FF9933&
            BorderColor     =   &H00FF9933&
            FillColor       =   &H00FF9933&
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox picSwitchErrors 
         Appearance      =   0  'Flat
         BackColor       =   &H00373333&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   60
         ScaleHeight     =   405
         ScaleWidth      =   1215
         TabIndex        =   3
         Top             =   60
         Width           =   1215
         Begin ImageX.aicAlphaImage imgError 
            Height          =   240
            Left            =   120
            Top             =   120
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            Image           =   "frmErrorList.frx":06F7
            Enabled         =   0   'False
            Props           =   13
         End
         Begin VB.Shape shpBorderErrors 
            BackColor       =   &H00FF9933&
            BorderColor     =   &H00FF9933&
            FillColor       =   &H00FF9933&
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label labErrorCount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 错误"
            ForeColor       =   &H00F0F0F0&
            Height          =   195
            Left            =   480
            TabIndex        =   4
            Top             =   120
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox picErrorType 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00373333&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   0
      ScaleHeight     =   2400
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   840
      Width           =   360
   End
   Begin DragControlsIDE.DarkListView lvErrorList 
      Height          =   2655
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4683
   End
   Begin VB.Image imgInfoIcon 
      Enabled         =   0   'False
      Height          =   240
      Left            =   5880
      Picture         =   "frmErrorList.frx":0882
      Top             =   3840
      Width           =   240
   End
   Begin VB.Image imgWarningIcon 
      Enabled         =   0   'False
      Height          =   240
      Left            =   5400
      Picture         =   "frmErrorList.frx":0C0C
      Top             =   3840
      Width           =   240
   End
   Begin VB.Image imgErrorIcon 
      Enabled         =   0   'False
      Height          =   240
      Left            =   4920
      Picture         =   "frmErrorList.frx":0F96
      Top             =   3840
      Width           =   240
   End
End
Attribute VB_Name = "frmErrorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'描述:      错误列表窗口，显示g++输出的错误和警告
'作者:      冰棍
'文件:      frmErrorList.frm
'====================================================

Option Explicit

'定义按钮颜色常数
Private Const NORMAL_COLOR = &H373333
Private Const MOUSEON_COLOR = &H5E5C5C
Private Const MOUSEDOWN_COLOR = &HCC7A00

'定义错误信息结构
Private Type ErrorInfo
    ErrorType                   As Byte                                     '错误类型（0: error; 1: warning; 2: info）
    Description                 As String                                   '描述
    FileName                    As String                                   '文件名
    FileLine                    As Long                                     '对应行
    FileColumn                  As Long                                     '对应列
End Type

Dim ErrorInfoList()             As ErrorInfo                                '所有错误信息（最后一个元素是多余的）
Dim InfoIndexBindingList()      As Long                                     'InfoIndexBindingList(列表项序号) = ErrorInfoList中的对应元素序号
Dim ColumnHeaderHeight          As Long                                     'ListView的ColumnHeader高度
Dim ListItemHeight              As Long                                     'ListView每个列表项的高度
Dim SpaceCount                  As Integer                                  '图片框的宽度相当于多少个空格
Dim ColumnHeader                As Long                                     '列表头的窗口句柄

Dim ShowErrors                  As Boolean                                  '用户自定义显示的消息类型
Dim ShowWarnings                As Boolean
Dim ShowInfo                    As Boolean
Dim ErrorCount                  As Long                                     '各种错误类型的计数
Dim WarningCount                As Long
Dim MessageCount                As Long

'描述:      清空所有东西，为下一次调试做准备
Public Sub ClearEverything()
    ErrorCount = 0
    WarningCount = 0
    MessageCount = 0
    Me.lvErrorList.Clear
    ReDim ErrorInfoList(0)
    ReDim InfoIndexBindingList(0)
    Me.picErrorType.Cls
    Call AddErrorListItem
End Sub

'描述:      重绘所有的错误类型图标
Public Sub RedrawErrorIcons()
    On Error Resume Next
    Dim i                       As Long
    Dim CurrentListCount        As Long
    Dim TopItem                 As Long, BottomItem                 As Long
    
    Me.picErrorType.Cls
    TopItem = Me.lvErrorList.GetTopIndex()                                  '获取ListView中第一个可视的列表项的序号
    BottomItem = TopItem + Me.lvErrorList.Height / ListItemHeight           '计算出ListView中最后一个可视的列表项的序号
    CurrentListCount = Me.lvErrorList.GetItemCount - 1                      '获取列表项数目
    If BottomItem > CurrentListCount Then                                   '检测最后一个可视列表项的序号是否超出列表项索引
        BottomItem = CurrentListCount
    End If
    For i = TopItem To BottomItem                                           '绘画可视范围内的所有图标
        If ErrorInfoList(InfoIndexBindingList(i)).ErrorType = 0 Then            'error
            Me.picErrorType.PaintPicture Me.imgErrorIcon.Picture, 0, (i - TopItem) * ListItemHeight + 60
        ElseIf ErrorInfoList(InfoIndexBindingList(i)).ErrorType = 1 Then        'warning
            Me.picErrorType.PaintPicture Me.imgWarningIcon.Picture, 0, (i - TopItem) * ListItemHeight + 60
        Else                                                                    'info
            Me.picErrorType.PaintPicture Me.imgInfoIcon.Picture, 0, (i - TopItem) * ListItemHeight + 60
        End If
    Next i
End Sub

'描述:      往ErrorInfoList和ListView里添加项目
'参数:      ErrorType: 错误类型（0: error; 1: warning; 2: info）
'.          Description: 错误描述
'.          FileName: 文件名
'.          FileLine: 对应行
'.          FileColumn: 对应列
Public Sub AddErrorInfoListItem(ErrorType As Byte, Description As String, FileName As String, FileLine As Long, FileColumn As Long)
    Dim tmpInfo                 As ErrorInfo
    
    '记录信息并添加到ErrorInfoList数组中
    tmpInfo.ErrorType = ErrorType
    tmpInfo.Description = Description
    tmpInfo.FileName = FileName
    tmpInfo.FileLine = FileLine
    tmpInfo.FileColumn = FileColumn
    
    '更新错误类型计数
    If ErrorType = 0 Then
        ErrorCount = ErrorCount + 1
    ElseIf ErrorType = 1 Then
        WarningCount = WarningCount + 1
    Else
        MessageCount = MessageCount + 1
    End If
    
    ErrorInfoList(UBound(ErrorInfoList)) = tmpInfo
    ReDim Preserve ErrorInfoList(UBound(ErrorInfoList) + 1)
End Sub

'描述:      按照用户当前选择显示的消息类型来添加ListView列表项
Public Sub AddErrorListItem()
    Dim i                       As Long
    Dim NewItemIndex            As Long
    
    '更新错误类型按钮上的计数
    Me.labErrorCount.Caption = ErrorCount & Lang_ErrorList_Errors
    Me.labWarningCount.Caption = WarningCount & Lang_ErrorList_Warnings
    Me.labInfoCount.Caption = MessageCount & Lang_ErrorList_Info
    
    '重新排版错误类型按钮
    Me.picSwitchErrors.Width = Me.labErrorCount.Left + Me.labErrorCount.Width + 120
    Me.shpBorderErrors.Width = Me.picSwitchErrors.Width
    Me.picSwitchWarnings.Left = Me.picSwitchErrors.Left + Me.picSwitchErrors.Width + 120
    Me.picSwitchWarnings.Width = Me.labWarningCount.Left + Me.labWarningCount.Width + 120
    Me.shpBorderWarnings.Width = Me.picSwitchWarnings.Width
    Me.picSwitchInfo.Left = Me.picSwitchWarnings.Left + Me.picSwitchWarnings.Width + 120
    Me.picSwitchInfo.Width = Me.labInfoCount.Left + Me.labInfoCount.Width + 120
    Me.shpBorderInfo.Width = Me.picSwitchInfo.Width
    
    Me.lvErrorList.Clear
    If Not ShowErrors And Not ShowWarnings And Not ShowInfo Then                '用户选择啥都不显示
        Exit Sub
    End If
    
    ReDim InfoIndexBindingList(UBound(ErrorInfoList))                           '分配足够大的序号绑定列表
    For i = 0 To UBound(ErrorInfoList) - 1                                      '添加所有符合条件的错误输出到ListView里
        If (ErrorInfoList(i).ErrorType = 0 And ShowErrors) Or _
           (ErrorInfoList(i).ErrorType = 1 And ShowWarnings) Or _
           (ErrorInfoList(i).ErrorType = 2 And ShowInfo) Then
        
            NewItemIndex = Me.lvErrorList.AddItem(Space(SpaceCount) & CStr(i + 1))
            Me.lvErrorList.SetItemText ErrorInfoList(i).Description, NewItemIndex, 1
            Me.lvErrorList.SetItemText ErrorInfoList(i).FileName, NewItemIndex, 2
            Me.lvErrorList.SetItemText CStr(ErrorInfoList(i).FileLine), NewItemIndex, 3
            Me.lvErrorList.SetItemText CStr(ErrorInfoList(i).FileColumn), NewItemIndex, 4
            InfoIndexBindingList(NewItemIndex) = i                                      '添加到序号绑定列表
        End If
    Next i
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_ErrorList_Caption
    
    '初始化控件排版
    Me.picErrorType.Move 0, Me.picTypeSelection.Height, 300
    Me.lvErrorList.Move 0, Me.picTypeSelection.Height
    
    Me.picSwitchErrors.Top = Me.picTypeSelection.Height / 2 - Me.picSwitchErrors.Height / 2
    Me.picSwitchErrors.Left = 120
    Me.imgError.Left = 60
    Me.imgError.Top = Me.picSwitchErrors.Height / 2 - Me.imgError.Height / 2
    Me.labErrorCount.Left = Me.imgError.Left + Me.imgError.Width + 120
    Me.labErrorCount.Top = Me.picSwitchErrors.Height / 2 - Me.labErrorCount.Height / 2
    Me.shpBorderErrors.Height = Me.picSwitchErrors.Height
    
    Me.picSwitchWarnings.Top = Me.picTypeSelection.Height / 2 - Me.picSwitchWarnings.Height / 2
    Me.imgWarning.Left = 60
    Me.imgWarning.Top = Me.picSwitchWarnings.Height / 2 - Me.imgWarning.Height / 2
    Me.labWarningCount.Left = Me.imgWarning.Left + Me.imgWarning.Width + 120
    Me.labWarningCount.Top = Me.picSwitchWarnings.Height / 2 - Me.labWarningCount.Height / 2
    Me.shpBorderWarnings.Height = Me.picSwitchErrors.Height
    
    Me.picSwitchInfo.Top = Me.picTypeSelection.Height / 2 - Me.picSwitchInfo.Height / 2
    Me.imgInfo.Left = 60
    Me.imgInfo.Top = Me.picSwitchInfo.Height / 2 - Me.imgInfo.Height / 2
    Me.labInfoCount.Left = Me.imgInfo.Left + Me.imgInfo.Width + 120
    Me.labInfoCount.Top = Me.picSwitchInfo.Height / 2 - Me.labInfoCount.Height / 2
    Me.shpBorderInfo.Height = Me.picSwitchInfo.Height
    
    '初始化ListView表头
    Me.lvErrorList.AddColumnHeader "#", 50
    Me.lvErrorList.AddColumnHeader Lang_ErrorList_Description, 300
    Me.lvErrorList.AddColumnHeader Lang_ErrorList_File, 310
    Me.lvErrorList.AddColumnHeader Lang_ErrorList_Line, 50
    Me.lvErrorList.AddColumnHeader Lang_ErrorList_Column, 50
    
    ReDim ErrorInfoList(0)                                                                          '初始化ErrorInfoList数组
    ShowErrors = True                                                                               '默认显示所有消息类型
    ShowWarnings = True
    ShowInfo = True
    Call AddErrorListItem
    
    '获取列表头的高度
    Dim tmpRect                 As RECT
    ColumnHeader = SendMessageA(Me.lvErrorList.ListViewHwnd, LVM_GETHEADER, 0, 0)                   '获取列表头的句柄
    SendMessageA ColumnHeader, HDM_GETITEMRECT, ByVal 0, ByVal VarPtr(tmpRect)                      '获取列表头的大小
    ColumnHeaderHeight = (tmpRect.bottom - tmpRect.Top) * Screen.TwipsPerPixelY                     '计算出列表头的高度
    
    '获取ListView中每个列表项的高度
    ZeroMemory tmpRect, ByVal Len(tmpRect)
    tmpRect.Left = LVIR_BOUNDS                                                                      '根据文档，在发消息前tmpRect.Left须设置为LVIR_BOUNDS
    Me.lvErrorList.AddItem "Stay DETERMINED!"                                                       '添加一个列表项，以计算列表项高度
    SendMessageA Me.lvErrorList.ListViewHwnd, LVM_GETITEMRECT, ByVal 0, ByVal VarPtr(tmpRect)       '获取列表项的大小
    Me.lvErrorList.Clear                                                                            '清空列表项
    ListItemHeight = (tmpRect.bottom - tmpRect.Top) * Screen.TwipsPerPixelY                         '计算出列表项的高度
    
    '设置列表头调整大小的窗口消息处理 todo
    SetPropA ColumnHeader, "PrevWndProc", SetWindowLongA(ColumnHeader, GWL_WNDPROC, AddressOf ErrorListColumnHeaderLayoutProc)
    
    '设置ListView重绘节点图标的窗口消息处理 todo
    SetPropA Me.lvErrorList.ListViewHwnd, "PrevWndProc", SetWindowLongA(Me.lvErrorList.ListViewHwnd, GWL_WNDPROC, AddressOf ErrorListIconRedrawProc)
    
    '把图片框放到ListView里
    SetParent Me.picErrorType.hwnd, Me.lvErrorList.ListViewHwnd
    Me.picErrorType.Top = ColumnHeaderHeight
    
    '计算图片框的宽度相当于多少个空格
    SpaceCount = Me.picErrorType.Width / Me.picErrorType.TextWidth(" ") + 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '恢复列表头的窗口消息处理
    SetWindowLongA ColumnHeader, GWL_WNDPROC, GetPropA(ColumnHeader, "PrevWndProc")
    
    '恢复ListView的窗口消息处理
    SetWindowLongA Me.lvErrorList.ListViewHwnd, GWL_WNDPROC, GetPropA(Me.lvErrorList.ListViewHwnd, "PrevWndProc")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.picErrorType.Height = Me.ScaleHeight - ColumnHeaderHeight
    Me.lvErrorList.Width = Me.ScaleWidth
    Me.lvErrorList.Height = Me.ScaleHeight - Me.picTypeSelection.Height
End Sub

Private Sub labErrorCount_Click()
    Call picSwitchErrors_Click
End Sub

Private Sub labErrorCount_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSwitchErrors_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub labErrorCount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSwitchErrors_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub labErrorCount_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSwitchErrors_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub labInfoCount_Click()
    Call picSwitchInfo_Click
End Sub

Private Sub labInfoCount_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSwitchInfo_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub labInfoCount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSwitchInfo_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub labInfoCount_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSwitchInfo_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub labWarningCount_Click()
    Call picSwitchWarnings_Click
End Sub

Private Sub labWarningCount_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSwitchWarnings_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub labWarningCount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSwitchWarnings_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub labWarningCount_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSwitchWarnings_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picSwitchErrors_Click()
    ShowErrors = Not ShowErrors
    Me.shpBorderErrors.Visible = ShowErrors
    Call AddErrorListItem
End Sub

Private Sub picSwitchErrors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.picSwitchErrors.BackColor = MOUSEDOWN_COLOR
End Sub

Private Sub picSwitchErrors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.tmrRestoreColor.Enabled = True
    If Me.picSwitchErrors.BackColor <> MOUSEON_COLOR And Me.picSwitchErrors.BackColor <> MOUSEDOWN_COLOR Then
        Me.picSwitchErrors.BackColor = MOUSEON_COLOR
    End If
End Sub

Private Sub picSwitchErrors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.picSwitchErrors.BackColor = NORMAL_COLOR
End Sub

Private Sub picSwitchInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.picSwitchInfo.BackColor = MOUSEDOWN_COLOR
End Sub

Private Sub picSwitchInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.tmrRestoreColor.Enabled = True
    If Me.picSwitchInfo.BackColor <> MOUSEON_COLOR And Me.picSwitchInfo.BackColor <> MOUSEDOWN_COLOR Then
        Me.picSwitchInfo.BackColor = MOUSEON_COLOR
    End If
End Sub

Private Sub picSwitchInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.picSwitchInfo.BackColor = NORMAL_COLOR
End Sub

Private Sub picSwitchWarnings_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.picSwitchWarnings.BackColor = MOUSEDOWN_COLOR
End Sub

Private Sub picSwitchWarnings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.tmrRestoreColor.Enabled = True
    If Me.picSwitchWarnings.BackColor <> MOUSEON_COLOR And Me.picSwitchWarnings.BackColor <> MOUSEDOWN_COLOR Then
        Me.picSwitchWarnings.BackColor = MOUSEON_COLOR
    End If
End Sub

Private Sub picSwitchWarnings_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.picSwitchWarnings.BackColor = NORMAL_COLOR
End Sub

Private Sub picSwitchInfo_Click()
    ShowInfo = Not ShowInfo
    Me.shpBorderInfo.Visible = ShowInfo
    Call AddErrorListItem
End Sub

Private Sub picSwitchWarnings_Click()
    ShowWarnings = Not ShowWarnings
    Me.shpBorderWarnings.Visible = ShowWarnings
    Call AddErrorListItem
End Sub

Private Sub lvErrorList_Click(iItem As Long, iSubItem As Long, X As Long, Y As Long)
    With ErrorInfoList(InfoIndexBindingList(iItem))
        Dim IconType    As Long
        
        If .ErrorType = 0 Then
            IconType = TTI_ERROR
        ElseIf .ErrorType = 1 Then
            IconType = TTI_WARNING
        Else
            IconType = TTI_INFO
        End If
        CtlAddToolTip Me.lvErrorList.ListViewHwnd, Lang_ErrorList_Description & ": " & .Description & vbCrLf & _
            Lang_ErrorList_File & ": " & .FileName & ":" & CStr(.FileLine) & ":" & CStr(.FileColumn), _
            Lang_ErrorList_Tooltip_Title, IconType
    End With
End Sub

Private Sub lvErrorList_DoubleClick(iItem As Long, iSubItem As Long, X As Long, Y As Long)
    Dim NewCodeWindow As frmCodeWindow
    
    '没有选择有效的列表项
    If iItem = -1 Or iItem > UBound(ErrorInfoList) Then
        Exit Sub
    End If
    
    '切换到对应的代码窗口
    Set NewCodeWindow = frmMain.ShowCodeWindow(, ErrorInfoList(InfoIndexBindingList(iItem)).FileName)
    If NewCodeWindow Is Nothing Then
        NoSkinMsgBox Lang_Main_Debug_OpenSourceFailure & ErrorInfoList(InfoIndexBindingList(iItem)).FileName, _
            vbExclamation, Lang_Msgbox_Error
    Else
        NewCodeWindow.SyntaxEdit.CurrPos.Row = ErrorInfoList(InfoIndexBindingList(iItem)).FileLine
        NewCodeWindow.SyntaxEdit.CurrPos.Col = ErrorInfoList(InfoIndexBindingList(iItem)).FileColumn
    End If
End Sub

Private Sub tmrRestoreColor_Timer()
    Dim CurPos          As POINT
    Dim CurrWindow      As Long
    Dim NeedToDisable   As Boolean
    
    '按着左键则不恢复颜色
    If GetAsyncKeyState(VK_LBUTTON) <> 0 Then
        Exit Sub
    End If
    
    '当检测到鼠标不在按钮上的时候就恢复按钮颜色
    NeedToDisable = True
    GetCursorPos CurPos
    CurrWindow = WindowFromPoint(CurPos.X, CurPos.Y)
    
    If CurrWindow <> Me.picSwitchErrors.hwnd Then
        Me.picSwitchErrors.BackColor = NORMAL_COLOR
    Else
        NeedToDisable = False
    End If
    
    If CurrWindow <> Me.picSwitchWarnings.hwnd Then
        Me.picSwitchWarnings.BackColor = NORMAL_COLOR
    Else
        NeedToDisable = False
    End If
    
    If CurrWindow <> Me.picSwitchInfo.hwnd Then
        Me.picSwitchInfo.BackColor = NORMAL_COLOR
    Else
        NeedToDisable = False
    End If
    
    If NeedToDisable Then
        Me.tmrRestoreColor.Enabled = False
    End If
End Sub

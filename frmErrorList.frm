VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "image.ocx"
Begin VB.Form frmErrorList 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "�����б�"
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
      TabIndex        =   1
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
         TabIndex        =   6
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
            Caption         =   "0 ��Ϣ"
            ForeColor       =   &H00F0F0F0&
            Height          =   195
            Left            =   480
            TabIndex        =   7
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
         TabIndex        =   4
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
            Caption         =   "0 ����"
            ForeColor       =   &H00F0F0F0&
            Height          =   195
            Left            =   480
            TabIndex        =   5
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
         TabIndex        =   2
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
            Caption         =   "0 ����"
            ForeColor       =   &H00F0F0F0&
            Height          =   195
            Left            =   480
            TabIndex        =   3
            Top             =   120
            Width           =   495
         End
      End
   End
   Begin DragControlsIDE.DarkListView lvErrorList 
      Height          =   2655
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4683
   End
End
Attribute VB_Name = "frmErrorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'����:      �����б��ڣ���ʾg++����Ĵ���;���
'����:      ����
'�ļ�:      frmErrorList.frm
'====================================================

Option Explicit

Private Declare Function LoadImageA Lib "user32" (ByVal hInst As Long, ByVal name As Long, ByVal restype As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal fuLoad As Long) As Long

'���尴ť��ɫ����
Private Const NORMAL_COLOR = &H373333
Private Const MOUSEON_COLOR = &H5E5C5C
Private Const MOUSEDOWN_COLOR = &HCC7A00

'����ͼ����ԴID����
Private Const IDI_ERROR = 101
Private Const IDI_INFO = 102
Private Const IDI_WARNING = 103

'���������Ϣ�ṹ
Private Type ErrorInfo
    ErrorType                   As Byte                                     '�������ͣ�0: error; 1: warning; 2: info��
    Description                 As String                                   '����
    FileName                    As String                                   '�ļ���
    FileLine                    As Long                                     '��Ӧ��
    FileColumn                  As Long                                     '��Ӧ��
End Type

Dim ErrorInfoList()             As ErrorInfo                                '���д�����Ϣ�����һ��Ԫ���Ƕ���ģ�
Dim InfoIndexBindingList()      As Long                                     'InfoIndexBindingList(�б������) = ErrorInfoList�еĶ�ӦԪ�����
Dim hImageList                  As Long                                     'ImageList���

Dim ShowErrors                  As Boolean                                  '�û��Զ�����ʾ����Ϣ����
Dim ShowWarnings                As Boolean
Dim ShowInfo                    As Boolean
Dim ErrorCount                  As Long                                     '���ִ������͵ļ���
Dim WarningCount                As Long
Dim MessageCount                As Long

'����:      ������ж�����Ϊ��һ�ε�����׼��
Public Sub ClearEverything()
    ErrorCount = 0
    WarningCount = 0
    MessageCount = 0
    Me.lvErrorList.Clear
    ReDim ErrorInfoList(0)
    ReDim InfoIndexBindingList(0)
    Call AddErrorListItem
End Sub

'����:      ��ErrorInfoList�������Ŀ
'����:      ErrorType: �������ͣ�0: error; 1: warning; 2: info��
'.          Description: ��������
'.          FileName: �ļ���
'.          FileLine: ��Ӧ��
'.          FileColumn: ��Ӧ��
Public Sub AddErrorInfoListItem(ErrorType As Byte, Description As String, FileName As String, FileLine As Long, FileColumn As Long)
    Dim tmpInfo                 As ErrorInfo
    
    '��¼��Ϣ����ӵ�ErrorInfoList������
    tmpInfo.ErrorType = ErrorType
    tmpInfo.Description = Description
    tmpInfo.FileName = FileName
    tmpInfo.FileLine = FileLine
    tmpInfo.FileColumn = FileColumn
    
    '���´������ͼ���
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

'����:      �����û���ǰѡ����ʾ����Ϣ���������ListView�б���
Public Sub AddErrorListItem()
    Dim i                       As Long
    Dim NewItemIndex            As Long
    
    '���´������Ͱ�ť�ϵļ���
    Me.labErrorCount.Caption = ErrorCount & Lang_ErrorList_Errors
    Me.labWarningCount.Caption = WarningCount & Lang_ErrorList_Warnings
    Me.labInfoCount.Caption = MessageCount & Lang_ErrorList_Info
    
    '�����Ű�������Ͱ�ť
    Me.picSwitchErrors.Width = Me.labErrorCount.Left + Me.labErrorCount.Width + 120
    Me.shpBorderErrors.Width = Me.picSwitchErrors.Width
    Me.picSwitchWarnings.Left = Me.picSwitchErrors.Left + Me.picSwitchErrors.Width + 120
    Me.picSwitchWarnings.Width = Me.labWarningCount.Left + Me.labWarningCount.Width + 120
    Me.shpBorderWarnings.Width = Me.picSwitchWarnings.Width
    Me.picSwitchInfo.Left = Me.picSwitchWarnings.Left + Me.picSwitchWarnings.Width + 120
    Me.picSwitchInfo.Width = Me.labInfoCount.Left + Me.labInfoCount.Width + 120
    Me.shpBorderInfo.Width = Me.picSwitchInfo.Width
    
    Me.lvErrorList.Clear
    If Not ShowErrors And Not ShowWarnings And Not ShowInfo Then                '�û�ѡ��ɶ������ʾ
        Exit Sub
    End If
    
    ReDim InfoIndexBindingList(UBound(ErrorInfoList))                           '�����㹻�����Ű��б�
    For i = 0 To UBound(ErrorInfoList) - 1                                      '������з��������Ĵ��������ListView��
        If (ErrorInfoList(i).ErrorType = 0 And ShowErrors) Or _
           (ErrorInfoList(i).ErrorType = 1 And ShowWarnings) Or _
           (ErrorInfoList(i).ErrorType = 2 And ShowInfo) Then
        
            NewItemIndex = Me.lvErrorList.AddItem(CStr(i + 1))
            Me.lvErrorList.SetItemText ErrorInfoList(i).Description, NewItemIndex, 1
            Me.lvErrorList.SetItemText ErrorInfoList(i).FileName, NewItemIndex, 2
            Me.lvErrorList.SetItemText CStr(ErrorInfoList(i).FileLine), NewItemIndex, 3
            Me.lvErrorList.SetItemText CStr(ErrorInfoList(i).FileColumn), NewItemIndex, 4
            
            '�����б����ͼ��
            Dim lvi As LVITEM
            
            lvi.mask = LVIF_IMAGE
            lvi.iItem = NewItemIndex
            lvi.iSubItem = 0
            lvi.iImage = CLng(ErrorInfoList(i).ErrorType)
            SendMessageA Me.lvErrorList.ListViewHwnd, LVM_SETITEM, ByVal 0, ByVal VarPtr(lvi)
            
            InfoIndexBindingList(NewItemIndex) = i                                      '��ӵ���Ű��б�
        End If
    Next i
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_ErrorList_Caption
    
    '��ʼ���ؼ��Ű�
    Me.lvErrorList.Left = 0
    Me.lvErrorList.Top = Me.picTypeSelection.Height
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
    
    '��ʼ��ListView��ͷ
    Me.lvErrorList.AddColumnHeader "#", 50
    Me.lvErrorList.AddColumnHeader Lang_ErrorList_Description, 300
    Me.lvErrorList.AddColumnHeader Lang_ErrorList_File, 310
    Me.lvErrorList.AddColumnHeader Lang_ErrorList_Line, 50
    Me.lvErrorList.AddColumnHeader Lang_ErrorList_Column, 50
    
    ReDim ErrorInfoList(0)                                                                          '��ʼ��ErrorInfoList����
    ShowErrors = True                                                                               'Ĭ����ʾ������Ϣ����
    ShowWarnings = True
    ShowInfo = True
    Call AddErrorListItem                                                                           '��ʼ����ť����
        
    '����ListView��ͼ���б�
    hImageList = ImageList_Create(16, 16, ILC_MASK, 0, 0)
    Dim hIcon   As Long
    ImageList_AddIcon hImageList, LoadImageA(App.hInstance, IDI_ERROR, IMAGE_ICON, 0, 0, LR_DEFAULTSIZE)
    ImageList_AddIcon hImageList, LoadImageA(App.hInstance, IDI_WARNING, IMAGE_ICON, 0, 0, LR_DEFAULTSIZE)
    ImageList_AddIcon hImageList, LoadImageA(App.hInstance, IDI_INFO, IMAGE_ICON, 0, 0, LR_DEFAULTSIZE)
    
    '��ListView��ͼ���б�
    SendMessageA Me.lvErrorList.ListViewHwnd, LVM_SETIMAGELIST, ByVal LVSIL_SMALL, ByVal hImageList
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '�ͷ�ImageList
    ImageList_Destroy hImageList
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
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
    
    'û��ѡ����Ч���б���
    If iItem = -1 Or iItem > UBound(ErrorInfoList) Then
        Exit Sub
    End If
    
    '�л�����Ӧ�Ĵ��봰��
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
    
    '��������򲻻ָ���ɫ
    If GetAsyncKeyState(VK_LBUTTON) <> 0 Then
        Exit Sub
    End If
    
    '����⵽��겻�ڰ�ť�ϵ�ʱ��ͻָ���ť��ɫ
    NeedToDisable = True
    GetCursorPos CurPos
    CurrWindow = WindowFromPoint(CurPos.X, CurPos.Y)
    
    If CurrWindow <> Me.picSwitchErrors.hwnd Then
        If Me.picSwitchErrors.BackColor <> NORMAL_COLOR Then
            Me.picSwitchErrors.BackColor = NORMAL_COLOR
        End If
    Else
        NeedToDisable = False
    End If
    
    If CurrWindow <> Me.picSwitchWarnings.hwnd Then
        If Me.picSwitchWarnings.BackColor <> NORMAL_COLOR Then
            Me.picSwitchWarnings.BackColor = NORMAL_COLOR
        End If
    Else
        NeedToDisable = False
    End If
    
    If CurrWindow <> Me.picSwitchInfo.hwnd Then
        If Me.picSwitchInfo.BackColor <> NORMAL_COLOR Then
            Me.picSwitchInfo.BackColor = NORMAL_COLOR
        End If
    Else
        NeedToDisable = False
    End If
    
    If NeedToDisable Then
        Me.tmrRestoreColor.Enabled = False
    End If
End Sub

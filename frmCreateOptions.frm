VERSION 5.00
Begin VB.Form frmCreateOptions 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "�½���Ŀ"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   Icon            =   "frmCreateOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   6840
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      Sizable         =   0   'False
   End
   Begin DragControlsIDE.DarkButton cmdCancel 
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "ȡ��"
      HasBorder       =   0   'False
   End
   Begin DragControlsIDE.DarkButton cmdOK 
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "ȷ��"
      HasBorder       =   0   'False
   End
   Begin DragControlsIDE.DarkCheckBox chkMain 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "����д��main ���޲�����"
   End
   Begin DragControlsIDE.DarkButton cmdBrowse 
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "���..."
      HasBorder       =   0   'False
   End
   Begin DragControlsIDE.DarkEdit edProjectName 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "�¿հ�C++����"
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
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
      Caption         =   "�½���Ŀ"
      MaxButtonEnabled=   0   'False
      MinButtonEnabled=   0   'False
      MaxButtonVisible=   0   'False
      MinButtonVisible=   0   'False
      BindCaption     =   -1  'True
      Picture         =   "frmCreateOptions.frx":1BCC2
   End
   Begin DragControlsIDE.DarkEdit edPath 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "C:\Project"
   End
   Begin DragControlsIDE.DarkCheckBox chkWinMain 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "����д��WinMain"
   End
   Begin DragControlsIDE.DarkCheckBox chkIncludeStdio 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "#include <stdio.h>"
   End
   Begin DragControlsIDE.DarkCheckBox chkMainArgs 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "����д��main ���в�����"
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ�ļ���:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ����:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Width           =   765
   End
End
Attribute VB_Name = "frmCreateOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'����:      �½�ѡ��ڣ��û�������������Ŀ�����ơ�·����һЩѡ��
'����:      ����
'�ļ�:      frmCreateOptions.frm
'====================================================

Option Explicit

'��ʾ������ļ��С��Ի���
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

'��ȡ����Ŀ¼·��
Private Declare Function SHGetFolderPathA Lib "shell32.dll" (ByVal hWnd As Long, ByVal csidl As Long, ByVal hToken As Long, _
    ByVal dwFlags As Long, pszPath As Any) As Long
    
Public NewProjectType   As Integer                                          '��Ҫ�½�����Ŀ���ͣ����frmMain��ProjectType����˵����
Dim MyDocPathStr        As String                                           '���ҵ��ĵ���·��
Dim PathChanged         As Boolean                                          '�û��Ƿ���Ĺ�·�������û���Ĺ���·����������Ŀ���ƶ��仯

Private Sub chkMain_Click()
    Me.chkMainArgs.Value = False
    Me.chkWinMain.Value = False
End Sub

Private Sub chkMainArgs_Click()
    Me.chkMain.Value = False
    Me.chkWinMain.Value = False
End Sub

Private Sub chkWinMain_Click()
    Me.chkMain.Value = False
    Me.chkMainArgs.Value = False
End Sub

Private Sub cmdBrowse_Click()
    '��ʾ��ѡ���ļ��С��Ի���
    Dim bi      As BROWSEINFO
    Dim pidl    As Long
    Dim NewPath As String * MAX_PATH
    
    With bi
        .hWndOwner = Me.hWnd
        .pidlRoot = 0
        .lpszTitle = Lang_CreateOptions_BrowseCaption
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_USENEWUI
    End With
    frmMain.SkinFramework.AutoApplyNewThreads = False                           '����Ƥ��
    frmMain.SkinFramework.AutoApplyNewWindows = False
    pidl = SHBrowseForFolder(bi)
    frmMain.SkinFramework.AutoApplyNewThreads = True                            '����Ƥ��
    frmMain.SkinFramework.AutoApplyNewWindows = True
    If pidl <> 0 Then                                                           '����û�û��ȡ������
        If SHGetPathFromIDList(pidl, NewPath) Then
            Me.edPath.Text = Split(NewPath, vbNullChar)(0)
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
    '���·��
    Dim ProjPath    As String
    
    ProjPath = IIf(Right(Me.edPath.Text, 1) = "\", Me.edPath.Text, Me.edPath.Text & "\")                '���"\"��·��ĩβ
    If Dir(Me.edPath.Text, vbDirectory Or vbNormal Or vbReadOnly Or vbHidden Or vbSystem) = "" Then     '��⵽·��������
        MkDir Me.edPath.Text                                                                                '���Դ����ļ���
        If Err.Number <> 0 Then                                                                             '�����ļ���ʧ��
            NoSkinMsgBox Lang_CreateOptions_InvalidProjectPath, vbExclamation, Lang_Msgbox_Error
            Me.cmdBrowse.SetFocus
            Exit Sub
        End If
    Else                                                                                                '��⵽·������
        If (GetAttr(Me.edPath.Text) And vbDirectory) = 0 Then                                               'Ŀ��·�������ļ���
            NoSkinMsgBox Lang_CreateOptions_InvalidProjectPath, vbExclamation, Lang_Msgbox_Error
            Me.cmdBrowse.SetFocus
            Exit Sub
        End If
    End If
    
    '���Դ�����Ŀ�ļ�
    Dim ProjCppPath As String                                                                           '��Ŀ����cpp�ļ�
    
    ProjCppPath = ProjPath & Me.edProjectName.Text & ".cpp"
    If Dir(ProjCppPath, vbDirectory Or vbNormal Or vbReadOnly Or vbHidden Or vbSystem) <> "" Then
        If NoSkinMsgBox(Lang_CreateOptions_NameConflict_1 & ProjCppPath & Lang_CreateOptions_NameConflict_2, vbQuestion Or vbYesNo, Lang_Msgbox_Confirm) <> vbYes Then
            Exit Sub
        End If
    End If
    Open ProjCppPath For Binary As #1
        If Err.Number <> 0 Then                                                                             '�����ļ�ʧ��
            Close #1
            NoSkinMsgBox Lang_CreateOptions_CreationFailure_1 & ProjCppPath & " :" & Err.Number & " - " & Err.Description & Lang_CreateOptions_CreationFailure_2, vbExclamation, Lang_Msgbox_Error
            Me.edProjectName.SetFocus
            Exit Sub
        End If
    Close #1
    
    '���´���״̬
    frmMain.ProjectType = NewProjectType                                                                '���ù�������
    Call frmMain.HideStartupPage                                                                        '������������
    If NewProjectType = 2 Or NewProjectType = 3 Then                                                    '������Ǵ��ڳ���ͽ��ö�Ӧ�Ĳ˵�
        frmMain.DarkMenu.MenuEnabled(29) = False                                                            '���ÿؼ���˵�
        frmMain.DarkMenu.MenuEnabled(30) = False                                                            '�������Բ˵�
    End If
    frmMain.DockingPane.ShowPane 3                                                                      '��ʾ������Դ������
    frmMain.DockingPane.ShowPane 5                                                                      '��ʾ���
    frmMain.Caption = Me.edProjectName.Text & " - " & Lang_Application_Title                            '���ı���
    frmMain.SkinFramework.AutoApplyNewThreads = True                                                    '���¼���Ƥ������������Ĺ������Ͳ��ܻ�����
    frmMain.SkinFramework.AutoApplyNewWindows = True
    
    '�������̽ṹ
    Dim ParentItem      As Long                                                                         '����ͼ�ĸ��ڵ�
    Dim GenCode         As String                                                                       '���ɵĴ���
    Dim CodeStartLn     As Long                                                                         '���ɴ���������ڵ���
    Dim NewCodeWindow   As frmCodeWindow                                                                '�´����Ĵ��봰��
    
    CurrentProject.ProjectName = Me.edProjectName.Text
    frmSolutionExplorer.SolutionTreeView.RemoveItem 0                                                   '�������ͼ
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem(CurrentProject.ProjectName)               '�����Ŀ
    ProjectNameTvItem = ParentItem                                                                      '��¼�������ƶ�Ӧ������ͼ�б���
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem(Lang_CreateOptions_SourceFile, ParentItem)
    frmSolutionExplorer.SolutionTreeView.ExpandItems frmSolutionExplorer.SolutionTreeView.GetParentItem(ParentItem), 2
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem(Me.edProjectName.Text & ".cpp", ParentItem)
    frmSolutionExplorer.SolutionTreeView.ExpandItems frmSolutionExplorer.SolutionTreeView.GetParentItem(ParentItem), 2
    frmSolutionExplorer.SolutionTreeView.SelectItem ParentItem
    ReDim TvItemBinding(0)                                                                              '���һ������ͼ�б�����ļ���ŵİ�
    TvItemBinding(0).FileIndex = 0                                                                      '���ð�
    TvItemBinding(0).TVITEM = ParentItem
    
    CodeStartLn = 1
    If Me.chkIncludeStdio.Value = True Then                                                             '#include <stdio.h>
        GenCode = "#include <stdio.h>" & vbCrLf & vbCrLf
        CodeStartLn = CodeStartLn + 2
    End If
    If Me.chkMain.Value = True Then                                                                     'main (�޲���)
        'int main() {
        '[Tab]
        '}
        GenCode = GenCode & "int main() {" & vbCrLf & vbTab & vbCrLf & "}" & vbCrLf
        CodeStartLn = CodeStartLn + 1
    ElseIf Me.chkMainArgs.Value = True Then                                                             'main (�в���)
        'int main(int argc, char *argv[]) {
        '[Tab]
        '}
        GenCode = GenCode & "int main(int argc, char *argv[]) {" & vbCrLf & vbTab & vbCrLf & "}" & vbCrLf
         CodeStartLn = CodeStartLn + 1
    ElseIf Me.chkWinMain.Value = True Then                                                              'WinMain
        '#include <windows.h>
        '
        'int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
        '[Tab]
        '}
        GenCode = GenCode & "#include <windows.h>" & vbCrLf & vbCrLf & _
            "int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {" & vbCrLf & vbTab & vbCrLf & "}" & vbCrLf
         CodeStartLn = CodeStartLn + 3
    End If
    With CurrentProject                                                                                 '���ù�����Ϣ
        ReDim .Files(0)
        With .Files(0)
            .FilePath = ProjCppPath
            .Changed = True
            .IsHeaderFile = False
            .PrevLine = CodeStartLn
        End With
        .ProjectType = NewProjectType
        .ProjectName = Me.edProjectName.Text
        .Changed = True
    End With
    ProjectFolderPath = ProjPath                                                                        '������Ŀ�ļ���·��
    ProjectFilePath = ProjPath & Me.edProjectName.Text & ".myproj"                                      '������Ŀ�����ļ�·��
    Set NewCodeWindow = CreateNewCodeWindow(0)                                                          '�½�һ�����봰��
    NewCodeWindow.Caption = Me.edProjectName.Text & ".cpp"
    frmMain.TabBar.AddForm NewCodeWindow
    CodeWindows.Add NewCodeWindow
    NewCodeWindow.FileIndex = 0                                                                         '���ô��봰�ڶ�Ӧ�Ĵ����ļ����
    NewCodeWindow.SyntaxEdit.Text = GenCode
    frmMain.picWindowClientArea.Visible = True                                                          '��ʾ���ڿͻ���
    NewCodeWindow.SyntaxEdit.CurrPos.SetPos CodeStartLn, NewCodeWindow.SyntaxEdit.TabSize + 1           '���������ƶ����ʺϵ�λ��
    NewCodeWindow.SyntaxEdit.SetFocus                                                                   '�ô�����ý���
    Unload Me
End Sub

Private Sub edPath_Change()
    PathChanged = True                                                      '�û����и��Ĺ�·��
    Me.edPath.ToolTipText = Me.edPath.Text
End Sub

Private Sub edProjectName_Change()
    Me.edProjectName.ToolTipText = Me.edProjectName.Text
    If Not PathChanged Then                                                 '����û�û�и��Ĺ�·�������Զ�������Ŀ�ļ���·��
        Me.edPath.Text = MyDocPathStr & "\" & Me.edProjectName.Text
        PathChanged = False
    End If
End Sub

Private Sub edProjectName_GotFocus()
    Me.edProjectName.SelStart = 0
    Me.edProjectName.SelLength = Len(Me.edProjectName.Text)
End Sub

Private Sub edProjectName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then                                          '��Ӧ�س���
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then                                          '����Esc��ȡ���½�
        KeyAscii = 0
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_CreateOptions_Caption                                                 '���ô��ڱ���
    Me.labTip(0).Caption = Lang_CreateOptions_ProjectNameLabel
    Me.labTip(1).Caption = Lang_CreateOptions_ProjectFolderLabel
    Me.chkIncludeStdio.Caption = Lang_CreateOptions_Include
    Me.chkMain.Caption = Lang_CreateOptions_Main_NoArgs
    Me.chkMainArgs.Caption = Lang_CreateOptions_Main_Args
    Me.chkWinMain.Caption = Lang_CreateOptions_WinMain
    Me.cmdBrowse.Caption = Lang_CreateOptions_Browse
    Me.cmdCancel.Caption = Lang_CreateOptions_Cancel
    Me.cmdOK.Caption = Lang_CreateOptions_OK
    '---------------------------------------------------------------------
    
    frmMain.Enabled = False
    frmMain.DarkWindowBorderSizer.Bind = False
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    
    '��ȡ���ҵ��ĵ���·����ΪĬ��·��
    Dim MyDocPath(MAX_PATH) As Byte
    Dim rtn                 As Long
    
    rtn = SHGetFolderPathA(0, CSIDL_PERSONAL, 0, 0, MyDocPath(0))
    If rtn = S_OK Then
        MyDocPathStr = Split(StrConv(MyDocPath, vbUnicode), vbNullChar)(0) & "\MyProjects"
    End If
    
    '���ݲ�ͬ�Ĺ�������ȡ��ͬ������
    Select Case NewProjectType
    Case 1
        Me.edProjectName.Text = Lang_CreateOptions_WindowProgram
        Me.edPath.Text = MyDocPathStr & "\" & Lang_CreateOptions_WindowProgram
        
    Case 2
        Me.edProjectName.Text = Lang_CreateOptions_ConsoleProgram
        Me.edPath.Text = MyDocPathStr & "\" & Lang_CreateOptions_ConsoleProgram
        
    Case 3
        Me.edProjectName.Text = Lang_CreateOptions_PlainCPP
        Me.edPath.Text = MyDocPathStr & "\" & Lang_CreateOptions_PlainCPP
        
    End Select
    Me.edProjectName.ToolTipText = Me.edProjectName.Text
    Me.edPath.ToolTipText = Me.edPath.Text
    PathChanged = False                                                         '��¼Ϊ�û�û�и��Ĺ�·��
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmMain.ProjectType = 0 Then
        frmMain.ShowStartupPage
    End If
    frmMain.Enabled = True
    frmMain.DarkWindowBorderSizer.Bind = True
    frmMain.SkinFramework.AutoApplyNewThreads = True
    frmMain.SkinFramework.AutoApplyNewWindows = True
End Sub

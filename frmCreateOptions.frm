VERSION 5.00
Begin VB.Form frmCreateOptions 
   BackColor       =   &H00403D3D&
   BorderStyle     =   0  'None
   Caption         =   "�½���Ŀ"
   ClientHeight    =   6870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   Icon            =   "frmCreateOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmCreateOptions.frx":1BCC2
   ScaleHeight     =   6870
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin DragControlsIDE.ImgOptionBox TypeOption 
      Height          =   1380
      Index           =   1
      Left            =   384
      TabIndex        =   0
      Top             =   1392
      Width           =   1308
      _ExtentX        =   2302
      _ExtentY        =   2434
      Image           =   "frmCreateOptions.frx":1C42C
      Content         =   "���ڳ���"
   End
   Begin VB.PictureBox picBtnFrame 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00504D4D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   924
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   10110
      TabIndex        =   13
      Top             =   5940
      Width           =   10110
      Begin DragControlsIDE.DarkButton cmdCancel 
         Height          =   492
         Left            =   8568
         TabIndex        =   11
         Top             =   192
         Width           =   1332
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
         Height          =   492
         Left            =   6936
         TabIndex        =   10
         Top             =   192
         Width           =   1332
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
   End
   Begin VB.PictureBox picProjectFrame 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00302D2D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   5325
      ScaleHeight     =   5445
      ScaleWidth      =   4785
      TabIndex        =   14
      Top             =   495
      Width           =   4788
      Begin DragControlsIDE.DarkCheckBox chkMainArgs 
         Height          =   372
         Left            =   240
         TabIndex        =   7
         Top             =   3048
         Width           =   4215
         _ExtentX        =   7435
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
      Begin DragControlsIDE.DarkCheckBox chkMain 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2610
         Width           =   4215
         _ExtentX        =   7435
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
      Begin DragControlsIDE.DarkEdit edProjectName 
         Height          =   372
         Left            =   288
         TabIndex        =   3
         Top             =   1128
         Width           =   4140
         _ExtentX        =   7303
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
      Begin DragControlsIDE.DarkCheckBox chkIncludeStdio 
         Height          =   372
         Left            =   240
         TabIndex        =   9
         Top             =   3912
         Width           =   4215
         _ExtentX        =   7435
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
      Begin DragControlsIDE.DarkCheckBox chkWinMain 
         Height          =   372
         Left            =   240
         TabIndex        =   8
         Top             =   3480
         Width           =   4215
         _ExtentX        =   7435
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
      Begin DragControlsIDE.DarkButton cmdBrowse 
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   1920
         Width           =   870
         _ExtentX        =   1535
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
      Begin DragControlsIDE.DarkEdit edPath 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   3300
         _ExtentX        =   5821
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
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   324
         Index           =   2
         Left            =   288
         TabIndex        =   17
         Top             =   240
         Width           =   480
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ����:"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Index           =   0
         Left            =   288
         TabIndex        =   16
         Top             =   768
         Width           =   768
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ�ļ���:"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Index           =   1
         Left            =   288
         TabIndex        =   15
         Top             =   1608
         Width           =   948
      End
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   5355
      Top             =   690
      _ExtentX        =   847
      _ExtentY        =   847
      Sizable         =   0   'False
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar_NoDrop 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
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
      Picture         =   "frmCreateOptions.frx":1C5A8
   End
   Begin DragControlsIDE.ImgOptionBox TypeOption 
      Height          =   1380
      Index           =   2
      Left            =   2016
      TabIndex        =   1
      Top             =   1392
      Width           =   1308
      _ExtentX        =   2302
      _ExtentY        =   2434
      Image           =   "frmCreateOptions.frx":1CD22
      Content         =   "����̨����"
   End
   Begin DragControlsIDE.ImgOptionBox TypeOption 
      Height          =   1380
      Index           =   3
      Left            =   3648
      TabIndex        =   2
      Top             =   1392
      Width           =   1308
      _ExtentX        =   2302
      _ExtentY        =   2434
      Image           =   "frmCreateOptions.frx":1CED2
      Content         =   "�հ�C++����"
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   390
      TabIndex        =   18
      Top             =   750
      Width           =   960
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
Private Declare Function SHGetFolderPathA Lib "shell32.dll" (ByVal hwnd As Long, ByVal csidl As Long, ByVal hToken As Long, _
    ByVal dwFlags As Long, pszPath As Any) As Long
    
Public NewProjectType   As Integer                                          '��Ҫ�½�����Ŀ���ͣ����frmMain��ProjectType����˵����
Dim MyDocPathStr        As String                                           '���ҵ��ĵ���·��
Dim PathChanged         As Boolean                                          '�û��Ƿ���Ĺ�·�������û���Ĺ���·����������Ŀ���ƶ��仯

'����:      ���ݲ�ͬ�Ĺ�������ȡ��ͬ�����ֺͲ�ͬ��·����
Public Sub RefreshName()
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
End Sub

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
        .hWndOwner = Me.hwnd
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
    
    'ȥ��·�����������ߵĿո�
    Dim ProjPath        As String                                                                       '����·��
    Dim ProjName        As String                                                                       '��������
    Dim PrevPathChanged As Boolean                                                                      '֮ǰ��PathChanged��־
    
    PrevPathChanged = PathChanged
    Me.edPath.Text = Trim(Me.edPath.Text)
    Me.edProjectName.Text = Trim(Me.edProjectName.Text)
    PathChanged = PrevPathChanged
    ProjName = Me.edProjectName.Text
    ProjPath = Me.edPath.Text
    ProjPath = IIf(Right(ProjPath, 1) = "\", ProjPath, ProjPath & "\")                                  '���"\"��·��ĩβ
    If Len(Trim(ProjName)) = 0 Then                                                                     'û�����빤������
        NoSkinMsgBox Lang_CreateOptions_ProjectNameRequired, vbExclamation, Lang_Msgbox_Error
        Me.edProjectName.SetFocus
        Exit Sub
    End If
    If Not CheckInvalidFileName(ProjName) Then                                                          '��鹤�������Ƿ�Ƿ�
        NoSkinMsgBox Lang_CreateOptions_InvalidProjectName & ProjName, vbExclamation, Lang_Msgbox_Error
        Me.edProjectName.SetFocus
        Exit Sub
    End If
    
    '���·��
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
    
    '����Ƿ��������ļ�
    If Dir(ProjPath & ProjName & ".cpp", vbDirectory Or vbNormal Or vbReadOnly Or vbHidden Or vbSystem) <> "" Then      '��⵽������cpp�ļ�
        If NoSkinMsgBox(Lang_CreateOptions_SameNameReplace_1 & ProjPath & ProjName & ".cpp" & _
           Lang_CreateOptions_SameNameReplace_2, vbQuestion Or vbYesNo, Lang_Msgbox_Confirm) <> vbYes Then
            Exit Sub
        End If
    End If
    If Dir(ProjPath & ProjName & ".myproj", vbDirectory Or vbNormal Or vbReadOnly Or vbHidden Or vbSystem) <> "" Then   '��⵽�����Ĺ����ļ�
        If NoSkinMsgBox(Lang_CreateOptions_SameNameReplace_1 & ProjPath & ProjName & ".myproj" & _
           Lang_CreateOptions_SameNameReplace_2, vbQuestion Or vbYesNo, Lang_Msgbox_Confirm) <> vbYes Then
            Exit Sub
        End If
    End If
    
    '������Ŀ��ص��ļ�
    Kill ProjPath & ProjName & ".cpp"
    Kill ProjPath & ProjName & ".myproj"
    Err.Clear
    Open ProjPath & ProjName & ".myproj" For Binary As #1                                               '���������ļ�
        If Err.Number <> 0 Then
            MsgBox Lang_CreateOptions_CreationFailure_1 & ProjPath & ProjName & ".myproj" & _
                   Lang_CreateOptions_CreateProjectFailed, vbExclamation, Lang_Msgbox_Error
            Close #1
            Exit Sub
        End If
    Close #1
    Open ProjPath & ProjName & ".cpp" For Binary As #1                                                  '����cpp�ļ�
        If Err.Number <> 0 Then
            MsgBox Lang_CreateOptions_CreationFailure_1 & ProjPath & ProjName & ".cpp" & _
                   Lang_CreateOptions_CreateProjectFailed, vbExclamation, Lang_Msgbox_Error
            Close #1
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
    frmMain.Caption = ProjName & " - " & Lang_Application_Title                                         '���ı���
    frmMain.SkinFramework.AutoApplyNewThreads = True                                                    '���¼���Ƥ������������Ĺ������Ͳ��ܻ�����
    frmMain.SkinFramework.AutoApplyNewWindows = True
    
    '�������̽ṹ
    Dim ParentItem      As Long                                                                         '����ͼ�ĸ��ڵ�
    Dim GenCode         As String                                                                       '���ɵĴ���
    Dim CodeStartLn     As Long                                                                         '���ɴ���������ڵ���
    Dim NewCodeWindow   As frmCodeWindow                                                                '�´����Ĵ��봰��
    
    CurrentProject.ProjectName = ProjName
    frmSolutionExplorer.SolutionTreeView.RemoveItem 0                                                   '�������ͼ
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem(CurrentProject.ProjectName)               '�����Ŀ
    ProjectNameTvItem = ParentItem                                                                      '��¼�������ƶ�Ӧ������ͼ�б���
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem(ProjName & ".cpp", ParentItem)
    frmSolutionExplorer.SolutionTreeView.ExpandItems ProjectNameTvItem, 2
    frmSolutionExplorer.SolutionTreeView.SelectItem ParentItem
    ReDim TvItemBinding(0)                                                                              '���һ������ͼ�б�����ļ���ŵİ�
    TvItemBinding(0).FileIndex = 0                                                                      '���ð�
    TvItemBinding(0).TVITEM = ParentItem
    TvItemBinding(0).IsFolder = False
    
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
        ReDim .Folders(0)
        ReDim .Files(0)
        With .Files(0)
            .FilePath = ProjPath & ProjName & ".cpp"
            .FolderIndex = 0
            .Changed = True
            .PrevLine = CodeStartLn
            ReDim .Breakpoints(0)
        End With
        .ProjectType = NewProjectType
        .ProjectName = ProjName
        .Changed = True
    End With
    ProjectFolderPath = ProjPath                                                                        '������Ŀ�ļ���·��
    ProjectFilePath = ProjPath & ProjName & ".myproj"                                                   '������Ŀ�����ļ�·��
    Set NewCodeWindow = CreateNewCodeWindow(0)                                                          '�½�һ�����봰��
    NewCodeWindow.Caption = ProjName & ".cpp"
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

Private Sub edPath_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then                         '��ӦCtrl+A
        Me.edPath.SelStart = 0
        Me.edPath.SelLength = Len(Me.edPath.Text)
    End If
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

Private Sub edProjectName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then                         '��ӦCtrl+A
        Me.edProjectName.SelStart = 0
        Me.edProjectName.SelLength = Len(Me.edProjectName.Text)
    End If
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
    Me.Caption = Lang_CreateOptions_Caption                                 '���ô��ڱ���
    Me.labTip(0).Caption = Lang_CreateOptions_ProjectNameLabel
    Me.labTip(1).Caption = Lang_CreateOptions_ProjectFolderLabel
    Me.labTip(3).Caption = Lang_CreateOptions_ProjectType
    Me.TypeOption(1).Content = Lang_CreateOptions_TypeOption_1
    Me.TypeOption(2).Content = Lang_CreateOptions_TypeOption_2
    Me.TypeOption(3).Content = Lang_CreateOptions_TypeOption_3
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
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    
    '��ȡ���ҵ��ĵ���·����ΪĬ��·��
    Dim MyDocPath(MAX_PATH) As Byte
    Dim rtn                 As Long
    
    rtn = SHGetFolderPathA(0, CSIDL_PERSONAL, 0, 0, MyDocPath(0))
    If rtn = S_OK Then
        MyDocPathStr = ByteArrayConv(MyDocPath) & "\MyProjects"
    End If
    
    Call RefreshName
    PathChanged = False                                                         '��¼Ϊ�û�û�и��Ĺ�·��
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmMain.ProjectType = 0 Then
        frmMain.ShowStartupPage
    Else
        frmMain.HideStartupPage
    End If
    frmMain.Enabled = True
    frmMain.DarkWindowBorderSizer.Bind = True
    frmMain.SkinFramework.AutoApplyNewThreads = True
    frmMain.SkinFramework.AutoApplyNewWindows = True
End Sub

Private Sub TypeOption_Click(Index As Integer)
    On Error Resume Next
    NewProjectType = Index                                                      '������Ŀ����
    Call RefreshName
    Me.edProjectName.SetFocus
End Sub

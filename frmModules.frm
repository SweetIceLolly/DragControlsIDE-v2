VERSION 5.00
Begin VB.Form frmModules 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "ģ��"
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DragControlsIDE.DarkListView lvModules 
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
   End
End
Attribute VB_Name = "frmModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'����:      ģ�鴰�ڣ����ж�״̬����ʾ���Խ��̼��ص�ģ��
'����:      ����
'�ļ�:      frmModules.frm
'====================================================

Option Explicit

Dim LoadedModuleInfo()  As ModuleInfoStruct                 '�Ѽ���ģ�����Ϣ

'����:      ������ж�����Ϊ��һ�ε�����׼��
Public Sub ClearEverything()
    Me.lvModules.Clear
    ReDim LoadedModuleInfo(0)
End Sub

'����:      ��ȡģ���б�
Public Sub GetModules()
    'on error resume next
    Dim PipeOutput      As String                                       '�ܵ������
    Dim OutputLines()   As String                                       '�����ÿһ��
    Dim NewListItem     As Long                                         '����ӵ�ListView�б�������
    Dim rtnInfo         As ModuleInfoStruct                             '�����õ���ģ����Ϣ
    Dim i               As Long
    
    Me.lvModules.Clear
    frmMain.DockingPane.Panes(12).Title = Lang_Modules_Caption & Lang_DebugWindow_Retrieving_Caption
    
    frmMain.GdbPipe.ClearPipe                                           '��չܵ��������
    frmMain.GdbPipe.DosInput "info sharedlibrary" & vbCrLf              '��gdb���ͻ�ȡģ������
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) ", 2000                '��ȡgdb���
    
    OutputLines = Split(PipeOutput, vbCrLf)                             '���зָ���
    ReDim LoadedModuleInfo(UBound(OutputLines) - 2)                     '������Ϣ�б�Ԫ��
    For i = 1 To UBound(OutputLines) - 2                                '���н��з���
        If Trim(OutputLines(i)) <> "(gdb)" Then                             'ȥ�����������(gdb) ��
            rtnInfo = ParseModuleString(OutputLines(i))
            LoadedModuleInfo(i) = rtnInfo
            
            NewListItem = Me.lvModules.AddItem(CStr(i))                         '������б���
            Me.lvModules.SetItemText GetFileName(rtnInfo.File), NewListItem, 1
            Me.lvModules.SetItemText rtnInfo.File, NewListItem, 2
            Me.lvModules.SetItemText rtnInfo.From, NewListItem, 3
            Me.lvModules.SetItemText rtnInfo.To, NewListItem, 4
        End If
    Next i
    
    frmMain.DockingPane(12).Title = Lang_Modules_Caption
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_Modules_Caption
    
    Me.lvModules.Move 0, 0
    
    Me.lvModules.AddColumnHeader "#", 35
    Me.lvModules.AddColumnHeader Lang_Modules_ListViewHeader_FileName, 95
    Me.lvModules.AddColumnHeader Lang_Modules_ListViewHeader_FilePath, 270
    Me.lvModules.AddColumnHeader Lang_Modules_ListViewHeader_From, 90
    Me.lvModules.AddColumnHeader Lang_Modules_ListViewHeader_To, 90
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.lvModules.Width = Me.ScaleWidth
    Me.lvModules.Height = Me.ScaleHeight
End Sub

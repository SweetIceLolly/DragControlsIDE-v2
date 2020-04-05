VERSION 5.00
Begin VB.Form frmThreads 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "�߳�"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DragControlsIDE.DarkListView lvThreads 
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
   End
End
Attribute VB_Name = "frmThreads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'����:      �̴߳��ڣ����ж�״̬����ʾ���Խ��̵��߳�
'����:      ����
'�ļ�:      frmThreads.frm
'====================================================

Option Explicit

Dim ThreadInfo()        As ThreadInfoStruct

'����:      ������ж�����Ϊ��һ�ε�����׼��
Public Sub ClearEverything()
    Me.lvThreads.Clear
    ReDim ThreadInfo(0)
End Sub

'����:      ��ȡ�߳��б�
Public Sub GetThreads()
    'on error resume next
    Dim PipeOutput      As String                                       '�ܵ������
    Dim OutputLines()   As String                                       '�����ÿһ��
    Dim NewListItem     As Long                                         '����ӵ�ListView�б�������
    Dim rtnInfo         As ThreadInfoStruct                             '�����õ����߳���Ϣ
    Dim i               As Long
    
    Me.lvThreads.Clear
    frmMain.DockingPane.Panes(12).Title = Lang_Modules_Caption & Lang_DebugWindow_Retrieving_Caption
    
    frmMain.GdbPipe.ClearPipe                                           '��չܵ��������
    frmMain.GdbPipe.DosInput "info threads" & vbCrLf                    '��gdb���ͻ�ȡ�߳�����
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) ", 2000                '��ȡgdb���
    
    OutputLines = Split(PipeOutput, vbCrLf)                             '���зָ���
    ReDim ThreadInfo(UBound(OutputLines) - 1)                           '������Ϣ�б�Ԫ��
    For i = 1 To UBound(OutputLines) - 1                                '���н��з���
        If Trim(OutputLines(i)) <> "(gdb)" Then                             'ȥ�����������(gdb) ��
            rtnInfo = ParseThreadString(OutputLines(i))
            ThreadInfo(i) = rtnInfo
            
            NewListItem = Me.lvThreads.AddItem(CStr(i))                         '������б���
            Me.lvThreads.SetItemText GetFileName(rtnInfo.Id), NewListItem, 1
            Me.lvThreads.SetItemText rtnInfo.Frame, NewListItem, 2
            
        End If
    Next i
    
    frmMain.DockingPane(11).Title = Lang_Modules_Caption
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_Threads_Caption
    
    Me.lvThreads.Move 0, 0
    
    Me.lvThreads.AddColumnHeader "#", 35
    Me.lvThreads.AddColumnHeader "ID", 90
    Me.lvThreads.AddColumnHeader Lang_Breakpoints_ListViewHeader_Address, 350
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.lvThreads.Width = Me.ScaleWidth
    Me.lvThreads.Height = Me.ScaleHeight
End Sub

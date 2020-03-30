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
      _extentx        =   6376
      _extenty        =   4683
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
    Dim NewListItem         As Long                                         '����ӵ�ListView�б�������
    Dim rtnInfo             As CallStackInfoStruct                          '�����õ��ĵ��ö�ջ��Ϣ
    Dim i                   As Long
    
    Me.lvCallStack.Clear
    frmMain.DockingPane.Panes(10).Title = Lang_CallStack_Retrieving_Caption
    
    frmMain.GdbPipe.ClearPipe                                               '��չܵ��������
    frmMain.GdbPipe.DosInput "info stack" & vbCrLf                          '��gdb���ͻ�ȡ���ö�ջ����
    frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                          '��ȡgdb���
    
    OutputLines = Split(PipeOutput, vbCrLf)                                 '���зָ���
    ReDim CallStackInfo(UBound(OutputLines) - 1)                            '������Ϣ�б�Ԫ��
    For i = 0 To UBound(OutputLines)                                        '���н��з���
        If Trim(OutputLines(i)) <> "(gdb)" Then                                 'ȥ�����������(gdb) ��
            If Mid(OutputLines(i), Len(OutputLines(i))) = vbCr Then                 'ȥ���ַ�����β�Ļ��з�
                OutputLines(i) = Left(OutputLines(i), Len(OutputLines(i)) - 1)
            End If
            
            rtnInfo = ParseCallStackString(OutputLines(i))
            CallStackInfo(i) = rtnInfo
            
            NewListItem = Me.lvCallStack.AddItem(rtnInfo.Address)                   '������б���
            If rtnInfo.Args <> "" Then
                Me.lvCallStack.SetItemText rtnInfo.Args, NewListItem, 1
            Else
                rtnInfo.Args = Lang_CallStack_NoArg
                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 1
            End If
            If rtnInfo.File <> "" Then
                Me.lvCallStack.SetItemText rtnInfo.File, NewListItem, 2
            Else
                rtnInfo.File = Lang_CallStack_NoArg
                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 2
            End If
            If rtnInfo.Line <> 0 Then
                Me.lvCallStack.SetItemText CStr(rtnInfo.Line), NewListItem, 3
            Else
                Me.lvCallStack.SetItemText Lang_CallStack_NoArg, NewListItem, 3
            End If
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
    
    If CallStackInfo(iItem).File <> "" Then                                                 '����ж�Ӧ���ļ�
        Dim NewCodeWindow   As frmCodeWindow
        
        '�л�����Ӧ�Ĵ���
        Set NewCodeWindow = frmMain.ShowCodeWindow(, CallStackInfo(iItem).File)
        If NewCodeWindow Is Nothing Then
            NoSkinMsgBox Lang_Main_Debug_OpenSourceFailure & CallStackInfo(iItem).File, vbExclamation, Lang_Msgbox_Error
        Else
            NewCodeWindow.SyntaxEdit.CurrPos.Row = CallStackInfo(iItem).Line
            NewCodeWindow.SyntaxEdit.SetFocus
        End If
    End If
End Sub


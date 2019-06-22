VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "gdb�ܵ����Բ���"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStopRecvOutput 
      Caption         =   "ֹͣ��ȡ���"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdClosePipe 
      Caption         =   "�رչܵ�"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdBreak 
      Caption         =   "�жϳ���"
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox edCommand 
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdSendCommand 
      Caption         =   "��������"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox edOutput 
      Height          =   5415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   9975
   End
   Begin VB.CommandButton cmdStartPipe 
      Caption         =   "�����ܵ�"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'����:      ʾ�����룬�����ܵ���������gdb����TestRes.exe
'����:      ����
'�ļ�:      frmMain.frm
'====================================================

Option Explicit

'����������Ϣ
Private Type STARTUPINFO
    cb                      As Long
    lpReserved              As Long
    lpDesktop               As Long
    lpTitle                 As Long
    dwX                     As Long
    dwY                     As Long
    dwXSize                 As Long
    dwYSize                 As Long
    dwXCountChars           As Long
    dwYCountChars           As Long
    dwFillAttribute         As Long
    dwFlags                 As Long
    wShowWindow             As Integer
    cbReserved2             As Integer
    lpReserved2             As Long
    hStdInput               As Long
    hStdOutput              As Long
    hStdError               As Long
End Type

'������Ϣ
Private Type PROCESS_INFORMATION
    hProcess                As Long
    hThread                 As Long
    dwProcessId             As Long
    dwThreadID              As Long
End Type

'��ȫ����
Private Type SECURITY_ATTRIBUTES
    nLength                 As Long
    lpSecurityDescriptor    As Long
    bInheritHandle          As Long
End Type

'�ж�ָ���Ľ���
Private Declare Function DebugBreakProcess Lib "kernel32" (ByVal hProcess As Long) As Long

'�رվ��
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

'�򿪽���
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'��������
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'ɱ������
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

'OpenProcess, dwDesiredAccess
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF                 '���̴�Ȩ��

'CreateProcess, dwCreationFlags
Private Const NORMAL_PRIORITY_CLASS = &H20&                         '��ͨ���ȼ�
Private Const CREATE_SUSPENDED = &H4                                '����֮����������

Dim pi                          As PROCESS_INFORMATION              '������Ϣ

Private WithEvents GdbPipe      As clsPipe
Attribute GdbPipe.VB_VarHelpID = -1

Private Sub cmdClosePipe_Click()
    GdbPipe.CloseDosIO
End Sub

Private Sub cmdStopRecvOutput_Click()
    GdbPipe.StopRecvOutput
End Sub

Private Sub Form_Load()
    Set GdbPipe = New clsPipe
End Sub

Private Sub cmdSendCommand_Click()
    Dim OutStr  As String
    If GdbPipe.DosInput(Me.edCommand.Text & vbCrLf) = 0 Then
        MsgBox "��������ʧ�ܣ�"
        Exit Sub
    End If
    GdbPipe.DosOutput OutStr, "(gdb) "
    Me.edOutput.Text = OutStr
    Me.edCommand.Text = ""
    Me.edCommand.SetFocus
End Sub

Private Sub cmdBreak_Click()
    If pi.hProcess <> 0 Then
        If DebugBreakProcess(pi.hProcess) = 0 Then                  '�����ж�Ŀ�����
            MsgBox "����ָ������ʧ�ܣ�"
        Else
            Dim OutStr  As String
            
            GdbPipe.DosOutput OutStr, "(gdb) "
            Me.edOutput.Text = OutStr
        End If
    Else
        MsgBox "����ָ������ʧ�ܣ�"
    End If
End Sub

Private Sub cmdStartPipe_Click()
    Dim OutStr  As String
    Dim si      As STARTUPINFO                                      '����������Ϣ
    Dim sa      As SECURITY_ATTRIBUTES                              '��ȫ����
    
    '���������Խ���
    With sa                                                         '���ð�ȫ����
        .nLength = Len(sa)
        .bInheritHandle = 1                                             '����ɼ̳�
        .lpSecurityDescriptor = 0
    End With
    If CreateProcess(ByVal 0, "TestRes.exe", sa, sa, ByVal 1, NORMAL_PRIORITY_CLASS Or CREATE_SUSPENDED, ByVal 0, ByVal 0, si, pi) <> 1 Then
        MsgBox "��������ʧ�ܣ�"
        Exit Sub
    End If
    
    '����gdb�ܵ�
    If GdbPipe.InitDosIO("gdb -q -nw") = 0 Then
        MsgBox "�����ܵ�ʧ�ܣ�"
        Exit Sub
    End If
    
    GdbPipe.DosInput "attach " & pi.dwProcessId & vbCrLf            '���ӵ������Խ���
    GdbPipe.DosInput "continue" & vbCrLf                            'ʹĿ����̼�������
    
    GdbPipe.DosOutput OutStr, "(gdb) "
    Me.edOutput.Text = OutStr
    
    Me.Caption = "�ܵ���������gdb����ID: " & GdbPipe.dwProcessId & " TestRes.exe����ID: " & pi.dwProcessId
End Sub

Private Sub edCommand_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSendCommand_Click
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    TerminateProcess pi.hProcess, 0                                 '�ص�Ŀ�����
    Set GdbPipe = Nothing
End Sub

'����:      �ܵ������������
'����:      strOutput: �ܵ��е�����
Private Sub GdbPipe_Output(strOutput As String)
    Me.edOutput.Text = strOutput
    Me.Caption = "DosOutputִ����..."
End Sub

'����:      �ܵ�����������
Private Sub GdbPipe_OutputFinished()
    Me.Caption = "DosOutput�ѷ��أ�"
End Sub

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
Option Explicit

Private Declare Function DebugBreakProcess Lib "kernel32" (ByVal hProcess As Long) As Long
'�򿪽���
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'�رվ��
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF                  '���̴�Ȩ��

Private WithEvents GdbPipe As clsPipe
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

Public Function SuspendProcess(ProcessID As Long) As Long
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcessID)
    If hProcess <> 0 Then
        SuspendProcess = DebugBreakProcess(hProcess)
        CloseHandle hProcess
    Else
        SuspendProcess = 0
    End If
End Function

Private Sub cmdBreak_Click()
    On Error Resume Next
    MsgBox "������ܻ�ûд�ã�"
    If SuspendProcess(CLng(Me.edCommand.Text)) = 0 Then
        MsgBox "����ָ������ʧ�ܣ�"
    End If
End Sub

Private Sub cmdStartPipe_Click()
    Dim OutStr  As String
    
    If GdbPipe.InitDosIO("gdb -q -nw") = 0 Then
        MsgBox "�����ܵ�ʧ�ܣ�"
    End If
    
    GdbPipe.DosOutput OutStr, "(gdb) "
    Me.edOutput.Text = OutStr
End Sub

Private Sub edCommand_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSendCommand_Click
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set GdbPipe = Nothing
End Sub

Private Sub GdbPipe_Output(strOutput As String)
    Me.edOutput.Text = strOutput
    Me.Caption = "DosOutputִ����..."
End Sub

Private Sub GdbPipe_OutputFinished()
    Me.Caption = "DosOutput�ѷ��أ�"
End Sub

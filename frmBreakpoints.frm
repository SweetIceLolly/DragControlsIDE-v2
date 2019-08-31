VERSION 5.00
Begin VB.Form frmBreakpoints 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "�ϵ��б�"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   Icon            =   "frmBreakpoints.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DragControlsIDE.DarkListView lvBreakpoints 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      _ExtentX        =   8281
      _ExtentY        =   5318
      CheckBoxes      =   -1  'True
   End
End
Attribute VB_Name = "frmBreakpoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'����:      �ϵ㴰�ڣ�������ʾ�����ļ������жϵ㣬���ṩ����ϵ㡢���öϵ㡢ɾ���ϵ�Ȳ���
'����:      ����
'�ļ�:      frmBreakpoints.frm
'====================================================

Option Explicit

'����:      ������ж�����Ϊ��һ�ε�����׼��
Public Sub ClearEverything()
    Dim i                   As Long
    
    For i = 0 To Me.lvBreakpoints.GetItemCount                                                  '��նϵ��Ӧ�ĵ�ַ
        Me.lvBreakpoints.SetItemText "", i, 2
    Next i
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_Breakpoints_Caption
    
    SetWindowLongA Me.lvBreakpoints.ListViewHwnd, GWL_STYLE, _
        GetWindowLongA(Me.lvBreakpoints.ListViewHwnd, GWL_STYLE) And (Not LVS_SINGLESEL)        '��ListView֧�ֶ�ѡ
    Me.lvBreakpoints.AddColumnHeader Lang_Breakpoints_ListViewHeader_File, 150                  '���ListView��ͷ
    Me.lvBreakpoints.AddColumnHeader Lang_Breakpoints_ListViewHeader_Line
    Me.lvBreakpoints.AddColumnHeader Lang_Breakpoints_ListViewHeader_Address
End Sub

Private Sub Form_Resize()
    Me.lvBreakpoints.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub lvBreakpoints_Click(iItem As Long, iSubItem As Long, X As Long, Y As Long)
    'On Error Resume Next       'todo
    Dim strTip              As String                                                           '������ʾ�ı�
    Dim strAddress          As String                                                           '�ϵ��Ӧ�ĵ�ַ
    
    strTip = "�ϵ��� " & Me.lvBreakpoints.GetItemText(iItem) & ":" & Me.lvBreakpoints.GetItemText(iItem, 1) & vbCrLf
    strAddress = Me.lvBreakpoints.GetItemText(iItem, 2)
    If strAddress <> "" Then                                                                    '�б���������ʾ��Ӧ�ĵ�ַ
        strTip = strTip & "��Ӧ��ַ: " & strAddress & vbCrLf
    End If
    If Me.lvBreakpoints.GetItemChecked(iItem) Then                                              '�ϵ�������
        strTip = strTip & "(������)"
    Else                                                                                        '�ϵ��ѽ���
        strTip = strTip & "(�ѽ���)"
    End If
    CtlAddToolTip Me.lvBreakpoints.ListViewHwnd, strTip, "�ϵ���Ϣ", TTI_INFO
End Sub


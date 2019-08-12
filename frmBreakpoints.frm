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
      _extentx        =   8281
      _extenty        =   5318
      checkboxes      =   -1
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

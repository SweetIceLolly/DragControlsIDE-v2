VERSION 5.00
Begin VB.Form frmBreakpoints 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "断点列表"
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
'描述:      断点窗口，负责显示所有文件的所有断点，并提供激活断点、禁用断点、删除断点等操作
'作者:      冰棍
'文件:      frmBreakpoints.frm
'====================================================

Option Explicit

'描述:      清空所有东西，为下一次调试做准备
Public Sub ClearEverything()
    Dim i                   As Long
    
    For i = 0 To Me.lvBreakpoints.GetItemCount                                                  '清空断点对应的地址
        Me.lvBreakpoints.SetItemText "", i, 2
    Next i
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_Breakpoints_Caption
    
    SetWindowLongA Me.lvBreakpoints.ListViewHwnd, GWL_STYLE, _
        GetWindowLongA(Me.lvBreakpoints.ListViewHwnd, GWL_STYLE) And (Not LVS_SINGLESEL)        '让ListView支持多选
    Me.lvBreakpoints.AddColumnHeader Lang_Breakpoints_ListViewHeader_File, 150                  '添加ListView表头
    Me.lvBreakpoints.AddColumnHeader Lang_Breakpoints_ListViewHeader_Line
    Me.lvBreakpoints.AddColumnHeader Lang_Breakpoints_ListViewHeader_Address
End Sub

Private Sub Form_Resize()
    Me.lvBreakpoints.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub lvBreakpoints_Click(iItem As Long, iSubItem As Long, X As Long, Y As Long)
    'On Error Resume Next       'todo
    Dim strTip              As String                                                           '工具提示文本
    Dim strAddress          As String                                                           '断点对应的地址
    
    strTip = "断点于 " & Me.lvBreakpoints.GetItemText(iItem) & ":" & Me.lvBreakpoints.GetItemText(iItem, 1) & vbCrLf
    strAddress = Me.lvBreakpoints.GetItemText(iItem, 2)
    If strAddress <> "" Then                                                                    '列表项里有显示对应的地址
        strTip = strTip & "对应地址: " & strAddress & vbCrLf
    End If
    If Me.lvBreakpoints.GetItemChecked(iItem) Then                                              '断点已启用
        strTip = strTip & "(已启用)"
    Else                                                                                        '断点已禁用
        strTip = strTip & "(已禁用)"
    End If
    CtlAddToolTip Me.lvBreakpoints.ListViewHwnd, strTip, "断点信息", TTI_INFO
End Sub


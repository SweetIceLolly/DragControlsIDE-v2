VERSION 5.00
Begin VB.UserControl DarkListView 
   Appearance      =   0  'Flat
   BackColor       =   &H00423E3E&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   ScaleHeight     =   2640
   ScaleWidth      =   3540
   ToolboxBitmap   =   "DarkListView.ctx":0000
End
Attribute VB_Name = "DarkListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark♂ListView by IceLolly
'Date: 2018.8.30
'Huge modification made on 2019.7.21

Dim PrevColumnCount     As Integer

Dim lvHwnd              As Long
 
Event ItemSelectionChanged()
Event MouseMove(Button As Long, Shift As Long, X As Integer, Y As Integer)
Event MouseDown(Button As Integer, Shift As Long, X As Integer, Y As Integer)
Event MouseUp(Button As Integer, Shift As Long, X As Integer, Y As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event ColumnClick(HeaderIndex As Integer)
Event ListViewLostFocus()
Event ListViewGotFocus()
Event Click(iItem As Long, iSubItem As Long, X As Long, Y As Long)
Event DoubleClick(iItem As Long, iSubItem As Long, X As Long, Y As Long)

'Default Property Values:
Const m_def_FullRowSelect = True

'Property Variables:
Dim m_FullRowSelect As Boolean
Dim m_GridLines     As Boolean
Dim m_CheckBoxes    As Boolean

Dim CurrExStyle     As Long

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub RaiseMouseMove(Button As Long, Shift As Long, X As Integer, Y As Integer)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub RaiseMouseDown(Button As Integer, Shift As Long, X As Integer, Y As Integer)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub RaiseMouseUp(Button As Integer, Shift As Long, X As Integer, Y As Integer)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub RaiseKeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub RaiseKeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub RaiseItemSelectionChanged()
    RaiseEvent ItemSelectionChanged
End Sub

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub RaiseLostFocus()
    RaiseEvent ListViewLostFocus
End Sub

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub RaiseGotFocus()
    RaiseEvent ListViewGotFocus
End Sub

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub RaiseClick(iItem As Long, iSubItem As Long, X As Long, Y As Long)
    RaiseEvent Click(iItem, iSubItem, X, Y)
End Sub

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub RaiseDoubleClick(iItem As Long, iSubItem As Long, X As Long, Y As Long)
    RaiseEvent DoubleClick(iItem, iSubItem, X, Y)
End Sub

Public Function AddColumnHeader(Text As String, Optional Width As Integer = 75, Optional Index As Long = -1) As Long
    Dim lvCol       As LVCOLUMN
    Dim tmpStr()    As Byte
    
    tmpStr = StrConvEx(Text)
    With lvCol
        .mask = LVCF_WIDTH Or LVCF_TEXT Or LVCF_FMT
        .fmt = LVCFMT_LEFT
        .cx = Width
        .pszText = VarPtr(tmpStr(0))
        .cchTextMax = 225
    End With
    AddColumnHeader = SendMessageA(lvHwnd, LVM_INSERTCOLUMN, IIf(Index = -1, _
        SendMessageA(SendMessageA(lvHwnd, LVM_GETHEADER, 0, 0), HDM_GETITEMCOUNT, 0, 0), _
        Index), ByVal VarPtr(lvCol))
End Function

Public Function DeleteColumnHeader(Index As Long) As Long
    DeleteColumnHeader = SendMessageA(lvHwnd, LVM_DELETECOLUMN, Index, 0)
End Function

Public Function AddItem(Text As String, Optional Index As Long = -1) As Long
    Dim lvi         As LVITEM
    Dim tmpStr()    As Byte
    
    tmpStr = StrConvEx(Text)
    With lvi
        .iItem = IIf(Index = -1, SendMessageA(lvHwnd, LVM_GETITEMCOUNT, ByVal 0, ByVal 0), Index)
        .mask = LVIF_TEXT
        .pszText = VarPtr(tmpStr(0))
        .cchTextMax = 255
    End With
    AddItem = SendMessageA(lvHwnd, LVM_INSERTITEM, 0, ByVal VarPtr(lvi))
End Function

Public Function DeleteItem(Index As Long) As Long
    DeleteItem = SendMessageA(lvHwnd, LVM_DELETEITEM, Index, 0)
End Function

Public Function GetItemText(Index As Long, Optional SubItemIndex As Long = 0) As String
    Dim tmpStr(255) As Byte
    Dim lvi         As LVITEM
    
    With lvi
        .mask = LVIF_TEXT
        .cchTextMax = 255
        .pszText = VarPtr(tmpStr(0))
        .iItem = Index
        .iSubItem = SubItemIndex
    End With
    SendMessageA lvHwnd, LVM_GETITEM, 0, ByVal VarPtr(lvi)
    GetItemText = ByteArrayConv(tmpStr)
End Function

Public Function SetItemText(Text As String, Index As Long, Optional SubItemIndex As Long = 0) As Long
    Dim lvi         As LVITEM
    Dim tmpStr()    As Byte
    
    tmpStr = StrConvEx(Text)
    With lvi
        .iSubItem = SubItemIndex
        .mask = LVIF_TEXT
        .pszText = VarPtr(tmpStr(0))
        .cchTextMax = 255
    End With
    SetItemText = SendMessageA(lvHwnd, LVM_SETITEMTEXT, Index, ByVal VarPtr(lvi))
End Function

Public Function GetItemCount() As Long
    GetItemCount = SendMessageA(lvHwnd, LVM_GETITEMCOUNT, 0, 0)
End Function

Public Function GetColumnText(Index As Long) As String
    Dim tmpStr(255) As Byte
    Dim lvc         As LVCOLUMN
    
    With lvc
        .mask = LVCF_TEXT
        .cchTextMax = 255
        .pszText = VarPtr(tmpStr(0))
    End With
    SendMessageA lvHwnd, LVM_GETCOLUMN, Index, ByVal VarPtr(lvc)
    GetColumnText = ByteArrayConv(tmpStr)
End Function

Public Function SetColumnText(Index As Long, NewText As String) As Long
    Dim tmpStr()    As Byte
    Dim lvc         As LVCOLUMN
    
    tmpStr = StrConvEx(NewText)
    With lvc
        .mask = LVCF_TEXT
        .cchTextMax = 255
        .pszText = VarPtr(tmpStr(0))
    End With
    SetColumnText = SendMessageA(lvHwnd, LVM_SETCOLUMN, Index, ByVal VarPtr(lvc))
End Function

Public Function GetColumnWidth(Index As Long) As Long
    Dim lvc         As LVCOLUMN
    
    lvc.mask = LVCF_WIDTH
    SendMessageA lvHwnd, LVM_GETCOLUMN, Index, ByVal VarPtr(lvc)
    GetColumnWidth = lvc.cx
End Function

Public Function SetColumnWidth(Index As Long, NewWidth As Long) As Long
    SetColumnWidth = SendMessageA(lvHwnd, LVM_SETCOLUMNWIDTH, Index, ByVal NewWidth)
End Function

'描述:      设置列表项的勾选状态（只适用于有选择框的ListVIew）
'参数:      Index: 列表项序号
'.          bChecked: 勾选状态。True: 勾选; False: 不勾选
Public Sub SetItemChecked(Index As Long, bChecked As Boolean)
    Dim lvi         As LVITEM
    
    '资料: https://docs.microsoft.com/en-us/windows/win32/controls/lvm-setitemstate
    With lvi
        .stateMask = LVIS_STATEIMAGEMASK
        .state = IIf(bChecked, 2, 1) * (2 ^ 12)             'x * 2^12 = x << 12
    End With
    SendMessageA lvHwnd, LVM_SETITEMSTATE, ByVal Index, ByVal VarPtr(lvi)
End Sub

'描述:      获取列表项的勾选状态（只适用于有选择框的ListVIew）
'参数:      Index: 列表项序号
'返回值:    True: 勾选; False: 不勾选
Public Function GetItemChecked(Index As Long) As Boolean
    'x \ 2^12 = x >> 12
    GetItemChecked = ((SendMessageA(lvHwnd, LVM_GETITEMSTATE, ByVal Index, LVIS_STATEIMAGEMASK) \ (2 ^ 12) - 1) = 1)
End Function

Public Sub Clear()
    SendMessageA lvHwnd, LVM_DELETEALLITEMS, 0, 0
End Sub

Public Function EnsureVisible(Index As Long, bEnsure As Boolean) As Long
    EnsureVisible = SendMessageA(lvHwnd, LVM_ENSUREVISIBLE, Index, IIf(bEnsure, 1, 0))
End Function

Public Function FindItem(Text As String, Optional FullMatch As Boolean = True, Optional StartIndex As Long = -1) As Long
    Dim tmpStr()    As Byte
    Dim lvfi        As LVFINDINFO
    
    tmpStr = StrConvEx(Text)
    If Not FullMatch Then
        lvfi.Flags = LVFI_PARTIAL
    End If
    lvfi.Flags = lvfi.Flags Or LVFI_STRING
    lvfi.psz = VarPtr(tmpStr(0))
    FindItem = SendMessageA(lvHwnd, LVM_FINDITEM, StartIndex, ByVal VarPtr(lvfi))
End Function

Public Function SetTextColor(Color As Long) As Long
    SetTextColor = SendMessageA(lvHwnd, LVM_SETTEXTCOLOR, 0, Color)
End Function

Public Function GetTextColor() As Long
    GetTextColor = SendMessageA(lvHwnd, LVM_GETTEXTCOLOR, 0, 0)
End Function

Public Function Scroll(vScroll As Long, Optional hScroll As Long = 0)
    Scroll = SendMessageA(lvHwnd, LVM_SCROLL, hScroll, hScroll)
End Function

Public Function GetTopIndex() As Long
    GetTopIndex = SendMessageA(lvHwnd, LVM_GETTOPINDEX, 0, 0)
End Function

Public Function GetSelectedItem() As Long
    GetSelectedItem = SendMessageA(lvHwnd, LVM_GETNEXTITEM, -1, LVNI_SELECTED)
End Function

Public Function SetSelectedItem(Index As Long) As Long
    Dim lvi         As LVITEM
    
    With lvi
        .state = LVIS_FOCUSED Or LVIS_SELECTED
        .stateMask = &HF
    End With
    SetSelectedItem = SendMessageA(lvHwnd, LVM_SETITEMSTATE, Index, ByVal VarPtr(lvi))
End Function

Private Sub labColumnHeader_Click(Index As Integer)
    RaiseEvent ColumnClick(Index)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim li  As LVITEM
    
    If KeyCode = vbKeyUp Then
        li.iItem = SendMessageA(lvHwnd, LVM_GETNEXTITEM, -1, LVNI_SELECTED) - 1
        li.state = LVIS_FOCUSED Or LVIS_SELECTED
        li.mask = LVIF_STATE
        li.stateMask = LVIS_FOCUSED Or LVIS_SELECTED
        SendMessageA lvHwnd, LVM_SETITEMSTATE, li.iItem, ByVal VarPtr(li)
    ElseIf KeyCode = vbKeyDown Then
        li.iItem = SendMessageA(lvHwnd, LVM_GETNEXTITEM, -1, LVNI_SELECTED) + 1
        li.state = LVIS_FOCUSED Or LVIS_SELECTED
        li.mask = LVIF_STATE
        li.stateMask = LVIS_FOCUSED Or LVIS_SELECTED
        SendMessageA lvHwnd, LVM_SETITEMSTATE, li.iItem, ByVal VarPtr(li)
    End If
    EnsureVisible li.iItem, True
    EnsureVisible li.iItem, False
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Initialize()
    CurrExStyle = WS_EX_NOPARENTNOTIFY
    lvHwnd = CreateWindowExA(CurrExStyle, "SysListView32", "", _
        WS_VISIBLE Or WS_CHILD Or WS_TABSTOP Or LVS_ALIGNLEFT Or LVS_REPORT Or LVS_SINGLESEL, _
        0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, _
        UserControl.ScaleHeight / Screen.TwipsPerPixelY, UserControl.hWnd, 0, App.hInstance, 0)
    
    SetPropA lvHwnd, "ID", ByVal CtlListPushBack(Me)
    SetPropA lvHwnd, "PARENT_CTL", UserControl.hWnd
    
    SendMessageA lvHwnd, LVM_SETBKCOLOR, ByVal 0, ByVal RGB(51, 51, 55)
    SendMessageA lvHwnd, LVM_SETTEXTBKCOLOR, ByVal 0, ByVal RGB(51, 51, 55)
    SendMessageA lvHwnd, LVM_SETTEXTCOLOR, ByVal 0, ByVal RGB(240, 240, 240)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    If Ambient.UserMode Then
        PrevLVUserCtlProc = SetWindowLongA(UserControl.hWnd, GWL_WNDPROC, AddressOf ListViewNotifyMessageProc)
        PrevListViewProc = SetWindowLongA(lvHwnd, GWL_WNDPROC, AddressOf ListViewProc)
    End If
    
    m_FullRowSelect = PropBag.ReadProperty("FullRowSelect", m_def_FullRowSelect)
    If m_FullRowSelect Then
        CurrExStyle = CurrExStyle Or LVS_EX_FULLROWSELECT
    Else
        CurrExStyle = CurrExStyle And (Not LVS_EX_FULLROWSELECT)
    End If
    
    m_GridLines = PropBag.ReadProperty("GridLines", False)
    If m_GridLines Then
        CurrExStyle = CurrExStyle Or LVS_EX_GRIDLINES
    Else
        CurrExStyle = CurrExStyle And (Not LVS_EX_GRIDLINES)
    End If
    
    m_CheckBoxes = PropBag.ReadProperty("CheckBoxes", False)
    If m_CheckBoxes Then
        CurrExStyle = CurrExStyle Or LVS_EX_CHECKBOXES
    Else
        CurrExStyle = CurrExStyle And Not (LVS_EX_CHECKBOXES)
    End If
    
    SendMessageA lvHwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, CurrExStyle
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    SetWindowPos lvHwnd, 0, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY, SWP_NOZORDER
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Returns/Sets if whole row is selected when an item is selected"
    FullRowSelect = m_FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
    m_FullRowSelect = New_FullRowSelect
    If New_FullRowSelect Then
        CurrExStyle = CurrExStyle Or LVS_EX_FULLROWSELECT
    Else
        CurrExStyle = CurrExStyle And (Not LVS_EX_FULLROWSELECT)
    End If
    SendMessageA lvHwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, CurrExStyle
    PropertyChanged "FullRowSelect"
End Property

Public Property Get GridLines() As Boolean
    GridLines = m_GridLines
End Property

Public Property Let GridLines(ByVal New_GridLines As Boolean)
    m_GridLines = New_GridLines
    If New_GridLines Then
        CurrExStyle = CurrExStyle Or LVS_EX_GRIDLINES
    Else
        CurrExStyle = CurrExStyle And (Not LVS_EX_GRIDLINES)
    End If
    SendMessageA lvHwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, CurrExStyle
    PropertyChanged "GridLines"
End Property

Public Property Get CheckBoxes() As Boolean
    CheckBoxes = m_CheckBoxes
End Property

Public Property Let CheckBoxes(ByVal New_CheckBoxes As Boolean)
    m_CheckBoxes = New_CheckBoxes
    If New_CheckBoxes Then
        CurrExStyle = CurrExStyle Or LVS_EX_CHECKBOXES
    Else
        CurrExStyle = CurrExStyle And (Not LVS_EX_CHECKBOXES)
    End If
    SendMessageA lvHwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, CurrExStyle
    PropertyChanged "CheckBoxes"
End Property

Public Property Get ListViewHwnd() As Long
    ListViewHwnd = lvHwnd
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FullRowSelect = m_def_FullRowSelect
    m_GridLines = False
    m_CheckBoxes = False
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("FullRowSelect", m_FullRowSelect, m_def_FullRowSelect)
    Call PropBag.WriteProperty("GridLines", m_GridLines, False)
    Call PropBag.WriteProperty("CheckBoxes", m_CheckBoxes, False)
End Sub


Attribute VB_Name = "modListViewProc"
Option Explicit

'WndProc to handle messages for Dark♂ListView
'Date: 2018.8.30
'Huge modification made on 2019.7.21

Public PrevListViewProc     As Long
Public PrevLVUserCtlProc    As Long

Dim CtlList()               As DarkListView

Public Function CtlListPushBack(Ctl As DarkListView) As Integer
    On Error Resume Next
    Dim NewIndex        As Integer
    
    NewIndex = UBound(CtlList) + 1
    ReDim Preserve CtlList(NewIndex)
    If Err.Number <> 0 Then
        Err.Clear
        ReDim CtlList(0)
        NewIndex = 0
    End If
    Set CtlList(NewIndex) = Ctl
    CtlListPushBack = NewIndex
End Function

Public Function ListViewProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_DESTROY
            SetWindowLongA hWnd, GWL_WNDPROC, ByVal PrevListViewProc
        
        Case WM_MOUSEMOVE
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseMove wParam And Not (MK_CONTROL Or MK_SHIFT), _
                (wParam And MK_CONTROL) Or (wParam And MK_SHIFT), LoWord(lParam), HiWord(lParam)
        
        Case WM_LBUTTONDOWN
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseDown 1, GetShiftValue(wParam), LoWord(lParam), HiWord(lParam)
            
        Case WM_LBUTTONUP
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseUp 1, GetShiftValue(wParam), LoWord(lParam), HiWord(lParam)
        
        Case WM_RBUTTONDOWN
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseDown 2, GetShiftValue(wParam), LoWord(lParam), HiWord(lParam)
        
        Case WM_RBUTTONUP
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseUp 2, GetShiftValue(wParam), LoWord(lParam), HiWord(lParam)
        
        Case WM_MBUTTONDOWN
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseDown 4, GetShiftValue(wParam), LoWord(lParam), HiWord(lParam)
        
        Case WM_MBUTTONUP
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseUp 4, GetShiftValue(wParam), LoWord(lParam), HiWord(lParam)
        
        Case WM_SETFOCUS
            SetFocus GetPropA(hWnd, "PARENT_CTL")
        
        Case LVM_SETBKCOLOR, LVM_SETTEXTBKCOLOR                                     '拦截调整背景颜色消息，防止被皮肤控件更改颜色
            lParam = RGB(51, 51, 55)
        
        Case LVM_SETTEXTCOLOR                                                       '拦截调整文本颜色消息，防止被皮肤控件更改颜色
            lParam = RGB(240, 240, 240)
        
    End Select
    ListViewProc = CallWindowProc(PrevListViewProc, hWnd, uMsg, wParam, lParam)
End Function

Public Function ListViewNotifyMessageProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim nm              As NMHDR
    Dim nmlv            As NMLISTVIEW
    Dim nmcd            As NMLVCUSTOMDRAW
    Dim nmia            As NMITEMACTIVATE
    
    If uMsg = WM_NOTIFY Then
        CopyMemory nm, ByVal lParam, ByVal Len(nm)
        Select Case nm.code
            Case LVN_ITEMCHANGED
                CopyMemory nmlv, ByVal lParam, ByVal Len(nmlv)
                If nmlv.uOldState = LVIS_SELECTED Then
                    CtlList(GetPropA(nm.hWndFrom, "ID")).RaiseItemSelectionChanged
                End If

            Case NM_KILLFOCUS
                CtlList(GetPropA(nm.hWndFrom, "ID")).RaiseLostFocus

            Case NM_SETFOCUS
                CtlList(GetPropA(nm.hWndFrom, "ID")).RaiseGotFocus
            
            Case NM_CLICK                                                       '选择列表项
                CopyMemory nmia, ByVal lParam, Len(nmia)
                CtlList(GetPropA(nm.hWndFrom, "ID")).RaiseClick nmia.iItem, nmia.iSubItem, nmia.ptAction.X, nmia.ptAction.Y
            
            Case NM_DBLCLK                                                      '双击列表项
                CopyMemory nmia, ByVal lParam, Len(nmia)
                CtlList(GetPropA(nm.hWndFrom, "ID")).RaiseDoubleClick nmia.iItem, nmia.iSubItem, nmia.ptAction.X, nmia.ptAction.Y
            
        End Select
    End If
    ListViewNotifyMessageProc = CallWindowProc(PrevLVUserCtlProc, hWnd, uMsg, wParam, lParam)
End Function

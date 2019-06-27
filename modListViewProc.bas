Attribute VB_Name = "modListViewProc"
Option Explicit

'WndProc to handle messages for Dark°·ListView
'Date: 2018.8.30

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public PrevListViewProc     As Long
Public PrevLVUserCtlProc    As Long

Dim CtlList()               As DarkListView

Public Function MakeLong(wLow As Long, wHigh As Long) As Long
    MakeLong = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

Public Function HiWord(lValue As Long) As Integer
    If lValue And &H80000000 Then
        HiWord = (lValue \ 65535) - 1
    Else
        HiWord = lValue \ 65535
    End If
End Function
 
Public Function LoWord(lValue As Long) As Integer
    If lValue And &H8000& Then
        LoWord = &H8000 Or (lValue And &H7FFF&)
    Else
        LoWord = lValue And &HFFFF&
    End If
End Function

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
End Function

Public Function ListViewProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_DESTROY
            SetWindowLongA hWnd, GWL_WNDPROC, ByVal PrevListViewProc
        
        Case WM_MOUSEMOVE
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseMove wParam And Not (MK_CONTROL Or MK_SHIFT), _
                (wParam And MK_CONTROL) Or (wParam And MK_SHIFT), LoWord(lParam), HiWord(lParam)
        
        Case WM_LBUTTONDOWN
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseDown 1, (wParam And MK_CONTROL) Or (wParam And MK_SHIFT), LoWord(lParam), HiWord(lParam)
            
        Case WM_LBUTTONUP
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseUp 1, (wParam And MK_CONTROL) Or (wParam And MK_SHIFT), LoWord(lParam), HiWord(lParam)
        
        Case WM_RBUTTONDOWN
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseDown 2, (wParam And MK_CONTROL) Or (wParam And MK_SHIFT), LoWord(lParam), HiWord(lParam)
        
        Case WM_RBUTTONUP
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseUp 2, (wParam And MK_CONTROL) Or (wParam And MK_SHIFT), LoWord(lParam), HiWord(lParam)
        
        Case WM_MBUTTONDOWN
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseDown 4, (wParam And MK_CONTROL) Or (wParam And MK_SHIFT), LoWord(lParam), HiWord(lParam)
        
        Case WM_MBUTTONUP
            CtlList(GetPropA(hWnd, "ID")).RaiseMouseUp 4, (wParam And MK_CONTROL) Or (wParam And MK_SHIFT), LoWord(lParam), HiWord(lParam)
        
        Case WM_LBUTTONDBLCLK
            CtlList(GetPropA(hWnd, "ID")).RaiseDoubleClick 1, (wParam And MK_CONTROL) Or (wParam And MK_SHIFT), LoWord(lParam), HiWord(lParam)
        
        Case WM_RBUTTONDBLCLK
            CtlList(GetPropA(hWnd, "ID")).RaiseDoubleClick 2, (wParam And MK_CONTROL) Or (wParam And MK_SHIFT), LoWord(lParam), HiWord(lParam)
        
        Case WM_MBUTTONDBLCLK
            CtlList(GetPropA(hWnd, "ID")).RaiseDoubleClick 4, (wParam And MK_CONTROL) Or (wParam And MK_SHIFT), LoWord(lParam), HiWord(lParam)
        
        Case WM_SETFOCUS
            SetFocus GetPropA(hWnd, "PARENT_CTL")
            
    End Select
    ShowScrollBar hWnd, SB_BOTH, 0
    ListViewProc = CallWindowProc(PrevListViewProc, hWnd, uMsg, wParam, lParam)
End Function

Public Function ListViewNotifyMessageProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim nm              As NMHDR
    Dim nmlv            As NMLISTVIEW
    Dim nmcd            As NMLVCUSTOMDRAW
    
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
                
            Case NM_CUSTOMDRAW
                CopyMemory nmcd, ByVal lParam, ByVal Len(nmcd)
                Select Case nmcd.nmcd.dwDrawStage
                    Case CDDS_PREPAINT
                        ListViewNotifyMessageProc = CDRF_NOTIFYITEMDRAW
                        Exit Function
                    
                    Case CDDS_ITEMPREPAINT, CDDS_ITEMPREPAINT Or CDDS_SUBITEM
                        Dim hBrush      As Long
                        Dim ItemRect    As RECT
                        Dim tmpRect     As RECT
                        Dim tmpStr(255) As Byte
                        Dim lvi         As LVITEM
                        Dim hRgn        As Long
                        Dim i           As Long
                        Dim hHeader     As Long
                        
                        SendMessageA nm.hWndFrom, LVM_GETITEMRECT, nmcd.nmcd.dwItemSpec, ByVal VarPtr(ItemRect)
                        hRgn = CreateRectRgnIndirect(ItemRect)
                        SetTextColor nmcd.nmcd.hDC, RGB(240, 240, 240)
                        hHeader = SendMessageA(nm.hWndFrom, LVM_GETHEADER, 0, 0)
                        If nmcd.nmcd.dwItemSpec >= SendMessageA(nm.hWndFrom, LVM_GETTOPINDEX, 0, 0) Then
                            If SendMessageA(nm.hWndFrom, LVM_GETNEXTITEM, -1, LVNI_SELECTED Or LVIS_FOCUSED) = nmcd.nmcd.dwItemSpec Then
                                hBrush = CreateSolidBrush(RGB(71, 71, 75))
                                FillRgn nmcd.nmcd.hDC, hRgn, hBrush
                                DeleteObject hBrush
                                DeleteObject hRgn
                                For i = 0 To SendMessageA(hHeader, HDM_GETITEMCOUNT, 0, 0)
                                    With lvi
                                        .mask = LVIF_TEXT
                                        .cchTextMax = 255
                                        .pszText = VarPtr(tmpStr(0))
                                        .iItem = nmcd.nmcd.dwItemSpec
                                        .iSubItem = i
                                    End With
                                    SendMessageA nm.hWndFrom, LVM_GETITEM, 0, ByVal VarPtr(lvi)
                                    SendMessageA hHeader, HDM_GETITEMRECT, i, ByVal VarPtr(tmpRect)
                                    SetRect tmpRect, tmpRect.Left + ItemRect.Left, ItemRect.Top, tmpRect.Right, ItemRect.bottom
                                    DrawTextA nmcd.nmcd.hDC, Split(StrConv(tmpStr, vbUnicode), vbNullChar)(0), -1, _
                                        tmpRect, DT_LEFT Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
                                Next i
                            Else
                                hBrush = CreateSolidBrush(RGB(51, 51, 55))
                                FillRgn nmcd.nmcd.hDC, hRgn, hBrush
                                DeleteObject hBrush
                                DeleteObject hRgn
                                For i = 0 To SendMessageA(hHeader, HDM_GETITEMCOUNT, 0, 0)
                                    With lvi
                                        .mask = LVIF_TEXT
                                        .cchTextMax = 255
                                        .pszText = VarPtr(tmpStr(0))
                                        .iItem = nmcd.nmcd.dwItemSpec
                                        .iSubItem = i
                                    End With
                                    SendMessageA nm.hWndFrom, LVM_GETITEM, 0, ByVal VarPtr(lvi)
                                    SendMessageA hHeader, HDM_GETITEMRECT, i, ByVal VarPtr(tmpRect)
                                    SetRect tmpRect, tmpRect.Left + ItemRect.Left, ItemRect.Top, tmpRect.Right, ItemRect.bottom
                                    DrawTextA nmcd.nmcd.hDC, Split(StrConv(tmpStr, vbUnicode), vbNullChar)(0), -1, _
                                        tmpRect, DT_LEFT Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
                                Next i
                            End If
                        End If
                        
                        ListViewNotifyMessageProc = CDRF_SKIPDEFAULT
                        Exit Function
                End Select
        End Select
    ElseIf uMsg = WM_MOUSEWHEEL Then
        If wParam < 0 Then
            CtlList(GetPropA(hWnd, "ID")).ScrollDown
            CtlList(GetPropA(hWnd, "ID")).ScrollDown
            CtlList(GetPropA(hWnd, "ID")).ScrollDown
        Else
            CtlList(GetPropA(hWnd, "ID")).ScrollUp
            CtlList(GetPropA(hWnd, "ID")).ScrollUp
            CtlList(GetPropA(hWnd, "ID")).ScrollUp
        End If
    End If
    ListViewNotifyMessageProc = CallWindowProc(PrevLVUserCtlProc, hWnd, uMsg, wParam, lParam)
End Function

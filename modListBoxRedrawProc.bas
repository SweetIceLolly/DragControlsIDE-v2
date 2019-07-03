Attribute VB_Name = "modListBoxRedrawProc"
 Option Explicit

'WndProc to support redrawing list items for Dark°·ListBox
'Date: 2018.8.26

Public PrevUserCtlProc      As Long
Public PrevListBoxProc      As Long

Public Function ListBoxWheelFixProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_MOUSEWHEEL Then
        If wParam < 0 Then
            SendMessageA hWnd, WM_VSCROLL, MakeLong(SB_LINEDOWN, 0), 0
        Else
            SendMessageA hWnd, WM_VSCROLL, MakeLong(SB_LINEUP, 0), 0
        End If
    End If
    ShowScrollBar hWnd, SB_BOTH, 0
    ListBoxWheelFixProc = CallWindowProc(PrevListBoxProc, hWnd, uMsg, wParam, ByVal lParam)
End Function

Public Function ListBoxRedrawProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim tItem               As DRAWITEMSTRUCT
    Dim sBuff()             As Byte
    Dim sItem               As String
    Dim hBkBrush            As Long
    Dim wndListBox          As Long

    wndListBox = FindWindowExA(hWnd, 0, "ThunderListBox", vbNullString)
    If uMsg = WM_DRAWITEM Then
        CopyMemory tItem, ByVal lParam, ByVal Len(tItem)
        If tItem.CtlType = ODT_LISTBOX Then
            ReDim sBuff(SendMessageA(tItem.hWndItem, LB_GETTEXTLEN, tItem.itemID, 0))
            SendMessageA tItem.hWndItem, LB_GETTEXT, tItem.itemID, ByVal VarPtr(sBuff(0))
            sItem = StrConv(sBuff, vbUnicode) & vbNullChar
            If (tItem.itemState And ODS_FOCUS) Then
                hBkBrush = CreateSolidBrush(ByVal RGB(71, 71, 72))
                FillRect tItem.hDC, tItem.rcItem, hBkBrush
                SetBkColor tItem.hDC, ByVal RGB(71, 71, 72)
                SetTextColor tItem.hDC, ByVal RGB(240, 240, 240)
                TextOutA tItem.hDC, tItem.rcItem.Left, tItem.rcItem.Top, ByVal sItem, ByVal UBound(sBuff)
            Else
                hBkBrush = CreateSolidBrush(ByVal RGB(51, 51, 55))
                FillRect tItem.hDC, tItem.rcItem, hBkBrush
                SetBkColor tItem.hDC, ByVal RGB(51, 51, 55)
                SetTextColor tItem.hDC, ByVal RGB(240, 240, 240)
                TextOutA tItem.hDC, tItem.rcItem.Left, tItem.rcItem.Top, ByVal sItem, ByVal UBound(sBuff)
            End If
            DeleteObject hBkBrush
            ListBoxRedrawProc = 0
            ShowScrollBar wndListBox, SB_BOTH, 0
            Exit Function
        End If
    ElseIf uMsg = WM_DESTROY Then
        SetWindowLongA hWnd, GWL_WNDPROC, ByVal PrevUserCtlProc
    Else
        ShowScrollBar wndListBox, SB_BOTH, 0
    End If

    ListBoxRedrawProc = CallWindowProc(PrevUserCtlProc, hWnd, uMsg, wParam, ByVal lParam)
End Function

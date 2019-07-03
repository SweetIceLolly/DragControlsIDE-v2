Attribute VB_Name = "modListWindowProc"
Option Explicit

'WndProc to support mouse wheel for Dark°·ComboBoxListWindow
'Date: 2018.8.11

Public PrevListWindowProc   As Long

Public Function ListWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_MOUSEWHEEL Then
        If frmComboBoxListWindow.VscrollBar.Visible = True Then
            If wParam < 0 Then
                frmComboBoxListWindow.VscrollBar.Value = frmComboBoxListWindow.VscrollBar.Value + frmComboBoxListWindow.VscrollBar.SmallChange
            Else
                frmComboBoxListWindow.VscrollBar.Value = frmComboBoxListWindow.VscrollBar.Value - frmComboBoxListWindow.VscrollBar.SmallChange
            End If
        End If
    End If
    ListWindowProc = CallWindowProc(PrevListWindowProc, hWnd, uMsg, wParam, lParam)
End Function

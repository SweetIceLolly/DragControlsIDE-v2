Attribute VB_Name = "modListWindowProc"
Option Explicit

'WndProc to support mouse wheel for Dark°·ComboBoxListWindow
'Date: 2018.8.11

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
Public PrevListWindowProc   As Long

Public Function ListWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_MOUSEWHEEL Then
        If frmComboBoxListWindow.VscrollBar.Visible = True Then
            If wParam < 0 Then
                frmComboBoxListWindow.VscrollBar.Value = frmComboBoxListWindow.VscrollBar.Value + frmComboBoxListWindow.VscrollBar.SmallChange
            Else
                frmComboBoxListWindow.VscrollBar.Value = frmComboBoxListWindow.VscrollBar.Value - frmComboBoxListWindow.VscrollBar.SmallChange
            End If
        End If
    End If
    ListWindowProc = CallWindowProc(PrevListWindowProc, hwnd, uMsg, wParam, lParam)
End Function

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
   Begin VB.Timer tmrUpdateScrollBars 
      Interval        =   50
      Left            =   1920
      Top             =   600
   End
   Begin 拖控件大法UI.DarkHScrollBar HScrollBar 
      Height          =   255
      Left            =   120
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      SmallChange     =   20
      LargeChange     =   60
   End
   Begin 拖控件大法UI.DarkVScrollBar VScrollBar 
      Height          =   1815
      Left            =   3000
      Top             =   360
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      SmallChange     =   20
      LargeChange     =   60
   End
   Begin VB.Timer tmrChangeHeaderSize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   600
   End
   Begin VB.Timer tmrCheckFocusAndColumn 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   600
   End
   Begin VB.PictureBox picResizer 
      Appearance      =   0  'Flat
      BackColor       =   &H00423E3E&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1320
      MousePointer    =   9  'Size W E
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label labColumnHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H00423E3E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Header"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "DarkListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark♂ListView by IceLolly
'Date: 2018.8.30
'This ListView can't edit labels due to its poor programming... Sorry about that O(╥﹏╥)O

Dim PrevColumnCount     As Integer

Dim hWnd                As Long
Dim NeedToCheckFocus    As Boolean          'Don't use this property
Dim DraggingIndex       As Integer
Dim IsDragging          As Boolean
Dim PrevRect            As RECT
Dim PrevIndex           As Integer
Dim ReqW                As Long, ReqH               As Long
 
Event ItemSelectionChanged()
Event MouseMove(Button As Long, Shift As Long, X As Integer, Y As Integer)
Event MouseDown(Button As Integer, Shift As Long, X As Integer, Y As Integer)
Event MouseUp(Button As Integer, Shift As Long, X As Integer, Y As Integer)
Event DoubleClick(Button As Integer, Shift As Long, X As Integer, Y As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event ColumnClick(HeaderIndex As Integer)
Event ListViewLostFocus()
Event ListViewGotFocus()
'Default Property Values:
Const m_def_FullRowSelect = True
'Property Variables:
Dim m_FullRowSelect As Boolean

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub UpdateHeaders(Optional ForceRefresh As Boolean = False)
    Dim CurrColumnCount As Long
    Dim hWndHeader      As Long
    
    If ForceRefresh Then
        GoTo lblRefresh
    Else
        hWndHeader = SendMessageA(hWnd, LVM_GETHEADER, 0, 0)
        CurrColumnCount = SendMessageA(hWndHeader, HDM_GETITEMCOUNT, 0, 0)
        If CurrColumnCount <> PrevColumnCount Then
            PrevColumnCount = CurrColumnCount
            GoTo lblRefresh
        Else
            Exit Sub
        End If
    End If
    
lblRefresh:
    Dim i               As Long
    Dim hdi             As HDITEM
    Dim tmpText(255)    As Byte
    
    If ForceRefresh Then
        hWndHeader = SendMessageA(hWnd, LVM_GETHEADER, 0, 0)
        CurrColumnCount = SendMessageA(hWndHeader, HDM_GETITEMCOUNT, 0, 0)
    End If
    LockWindowUpdate UserControl.hWnd
    If CurrColumnCount <= 0 Then
        For i = 1 To UserControl.labColumnHeader.UBound
            Unload UserControl.labColumnHeader(i)
            Unload UserControl.picResizer(i)
        Next i
        UserControl.labColumnHeader(0).Visible = False
        UserControl.picResizer(0).Visible = False
    Else
        hdi.mask = HDI_TEXT Or HDI_WIDTH
        hdi.cchTextMax = 255
        hdi.pszText = VarPtr(tmpText(0))
        SendMessageA hWndHeader, HDM_GETITEMA, 0, ByVal VarPtr(hdi)
        UserControl.labColumnHeader(0).Visible = True
        UserControl.picResizer(0).Visible = True
        UserControl.labColumnHeader(0).Caption = StrConv(tmpText, vbUnicode)
        UserControl.labColumnHeader(0).Width = hdi.cxy * Screen.TwipsPerPixelX
        If CurrColumnCount > 1 Then
            For i = 1 To UserControl.labColumnHeader.UBound
                Unload UserControl.labColumnHeader(i)
                Unload UserControl.picResizer(i)
            Next i
            For i = 2 To CurrColumnCount
                hdi.mask = HDI_TEXT Or HDI_WIDTH
                hdi.cchTextMax = 255
                hdi.pszText = VarPtr(tmpText(0))
                SendMessageA hWndHeader, HDM_GETITEMA, i - 1, ByVal VarPtr(hdi)
                Load UserControl.labColumnHeader(i - 1)
                Load UserControl.picResizer(i - 1)
                UserControl.labColumnHeader(i - 1).Visible = True
                UserControl.picResizer(i - 1).Visible = True
                UserControl.labColumnHeader(i - 1).Caption = StrConv(tmpText, vbUnicode)
                UserControl.labColumnHeader(i - 1).Width = hdi.cxy * Screen.TwipsPerPixelX
            Next i
        End If
        Call UserControl_Resize
    End If
    LockWindowUpdate 0
End Sub

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
Public Sub RaiseDoubleClick(Button As Integer, Shift As Long, X As Integer, Y As Integer)
    RaiseEvent DoubleClick(Button, Shift, X, Y)
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

Public Function AddColumnHeader(Text As String, Optional Width As Integer = 75, Optional Index As Long = -1) As Long
    
    Dim lvCol       As LVCOLUMN
    Dim tmpStr()    As Byte
    
    tmpStr = StrConv(Text & vbNullChar, vbFromUnicode)
    With lvCol
        .mask = LVCF_WIDTH Or LVCF_TEXT Or LVCF_FMT
        .fmt = LVCFMT_LEFT
        .cx = Width
        .pszText = VarPtr(tmpStr(0))
        .cchTextMax = 225
    End With
    AddColumnHeader = SendMessageA(hWnd, LVM_INSERTCOLUMN, IIf(Index = -1, _
        SendMessageA(SendMessageA(hWnd, LVM_GETHEADER, 0, 0), HDM_GETITEMCOUNT, 0, 0), _
        Index), ByVal VarPtr(lvCol))
End Function

Public Function DeleteColumnHeader(Index As Long) As Long
    DeleteColumnHeader = SendMessageA(hWnd, LVM_DELETECOLUMN, Index, 0)
End Function

Public Function AddItem(Text As String, Optional Index As Long = -1) As Long
    Dim lvi         As LVITEM
    Dim tmpStr()    As Byte
    
    tmpStr = StrConv(Text & vbNullChar, vbFromUnicode)
    With lvi
        .iItem = IIf(Index = -1, SendMessageA(hWnd, LVM_GETITEMCOUNT, ByVal 0, ByVal 0), Index)
        .mask = LVIF_TEXT
        .pszText = VarPtr(tmpStr(0))
        .cchTextMax = 255
    End With
    AddItem = SendMessageA(hWnd, LVM_INSERTITEM, 0, ByVal VarPtr(lvi))
End Function

Public Function DeleteItem(Index As Long) As Long
    DeleteItem = SendMessageA(hWnd, LVM_DELETEITEM, Index, 0)
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
    SendMessageA hWnd, LVM_GETITEM, 0, ByVal VarPtr(lvi)
    GetItemText = Split(StrConv(tmpStr, vbUnicode), vbNullChar)(0)
End Function

Public Function SetItemText(Text As String, Index As Long, Optional SubItemIndex As Long = 0) As Long
    Dim lvi         As LVITEM
    Dim tmpStr()    As Byte
    
    tmpStr = StrConv(Text & vbNullChar, vbFromUnicode)
    With lvi
        .iSubItem = SubItemIndex
        .mask = LVIF_TEXT
        .pszText = VarPtr(tmpStr(0))
        .cchTextMax = 255
    End With
    SetItemText = SendMessageA(hWnd, LVM_SETITEMTEXT, Index, ByVal VarPtr(lvi))
End Function

Public Function GetItemCount() As Long
    GetItemCount = SendMessageA(hWnd, LVM_GETITEMCOUNT, 0, 0)
End Function

Public Function GetColumnText(Index As Long) As String
    Dim tmpStr(255) As Byte
    Dim lvc         As LVCOLUMN
    
    With lvc
        .mask = LVCF_TEXT
        .cchTextMax = 255
        .pszText = VarPtr(tmpStr(0))
    End With
    SendMessageA hWnd, LVM_GETCOLUMN, Index, ByVal VarPtr(lvc)
    GetColumnText = Split(StrConv(tmpStr, vbUnicode), vbNullChar)(0)
End Function

Public Function SetColumnText(Index As Long, NewText As String) As Long
    Dim tmpStr()    As Byte
    Dim lvc         As LVCOLUMN
    
    tmpStr = StrConv(NewText & vbNullChar, vbFromUnicode)
    With lvc
        .mask = LVCF_TEXT
        .cchTextMax = 255
        .pszText = VarPtr(tmpStr(0))
    End With
    SetColumnText = SendMessageA(hWnd, LVM_SETCOLUMN, Index, ByVal VarPtr(lvc))
    Call UpdateHeaders(True)
End Function

Public Function GetColumnWidth(Index As Long) As Long
    Dim lvc         As LVCOLUMN
    
    lvc.mask = LVCF_WIDTH
    SendMessageA hWnd, LVM_GETCOLUMN, Index, ByVal VarPtr(lvc)
    GetColumnWidth = lvc.cx
End Function

Public Function SetColumnWidth(Index As Long, NewWidth As Long) As Long
    SetColumnWidth = SendMessageA(hWnd, LVM_SETCOLUMNWIDTH, Index, ByVal NewWidth)
    Call UpdateHeaders(True)
End Function

Public Sub Clear()
    SendMessageA hWnd, LVM_DELETEALLITEMS, 0, 0
End Sub

Public Function EnsureVisible(Index As Long, bEnsure As Boolean) As Long
    EnsureVisible = SendMessageA(hWnd, LVM_ENSUREVISIBLE, Index, IIf(bEnsure, 1, 0))
End Function

Public Function FindItem(Text As String, Optional FullMatch As Boolean = True, Optional StartIndex As Long = -1) As Long
    Dim tmpStr()    As Byte
    Dim lvfi        As LVFINDINFO
    
    tmpStr = StrConv(Text & vbNullChar, vbFromUnicode)
    If Not FullMatch Then
        lvfi.Flags = LVFI_PARTIAL
    End If
    lvfi.Flags = lvfi.Flags Or LVFI_STRING
    lvfi.psz = VarPtr(tmpStr(0))
    FindItem = SendMessageA(hWnd, LVM_FINDITEM, StartIndex, ByVal VarPtr(lvfi))
End Function

Public Function SetTextColor(Color As Long) As Long
    SetTextColor = SendMessageA(hWnd, LVM_SETTEXTCOLOR, 0, Color)
End Function

Public Function GetTextColor() As Long
    GetTextColor = SendMessageA(hWnd, LVM_GETTEXTCOLOR, 0, 0)
End Function

Public Function Scroll(vScroll As Long, Optional hScroll As Long = 0)
    Scroll = SendMessageA(hWnd, LVM_SCROLL, hScroll, hScroll)
End Function

Public Function GetTopIndex() As Long
    GetTopIndex = SendMessageA(hWnd, LVM_GETTOPINDEX, 0, 0)
End Function

Public Function GetSelectedItem() As Long
    GetSelectedItem = SendMessageA(hWnd, LVM_GETNEXTITEM, -1, LVNI_SELECTED)
End Function

Public Function SetSelectedItem(Index As Long) As Long
    Dim lvi         As LVITEM
    
    With lvi
        .state = LVIS_FOCUSED Or LVIS_SELECTED
        .stateMask = &HF
    End With
    SetSelectedItem = SendMessageA(hWnd, LVM_SETITEMSTATE, Index, ByVal VarPtr(lvi))
End Function

Public Sub ScrollDown()
    UserControl.VscrollBar.Value = UserControl.VscrollBar.Value + UserControl.VscrollBar.SmallChange
End Sub

Public Sub ScrollUp()
    UserControl.VscrollBar.Value = UserControl.VscrollBar.Value - UserControl.VscrollBar.SmallChange
End Sub

Private Sub labColumnHeader_Click(Index As Integer)
    RaiseEvent ColumnClick(Index)
End Sub

Private Sub labColumnHeader_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i           As Integer
    
    If Not IsDragging And PrevIndex <> Index Then
        For i = 0 To UserControl.labColumnHeader.UBound
            UserControl.labColumnHeader(i).BackColor = RGB(62, 62, 66)
        Next i
        UserControl.labColumnHeader(Index).BackColor = RGB(82, 82, 86)
        NeedToCheckFocus = True
        PrevIndex = Index
    End If
End Sub

Private Sub picResizer_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub picResizer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If Index > 0 Then
            UserControl.MousePointer = 9
            GetWindowRect UserControl.picResizer(Index - 1).hWnd, PrevRect
        Else
            GetWindowRect UserControl.hWnd, PrevRect
            PrevRect.Left = PrevRect.Left + UserControl.labColumnHeader(0).Left / Screen.TwipsPerPixelX
        End If
        IsDragging = True
        DraggingIndex = Index
        UserControl.tmrChangeHeaderSize.Enabled = True
    End If
End Sub

Private Sub picResizer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    IsDragging = False
    UserControl.tmrChangeHeaderSize.Enabled = False
End Sub

Private Sub tmrChangeHeaderSize_Timer()
    Static PrevNewW As Long
    
    If GetAsyncKeyState(VK_LBUTTON) = 0 Then
        If DraggingIndex > 0 Then
            UserControl.MousePointer = 0
        End If
        PrevNewW = -1
        IsDragging = False
        UserControl.tmrChangeHeaderSize.Enabled = False
        Exit Sub
    End If
    If IsDragging Then
        Dim cur         As POINT
        Dim NewWidth    As Long
        
        GetCursorPos cur
        If cur.X < PrevRect.Left + 6 Then
            NewWidth = 6
        Else
            NewWidth = cur.X - PrevRect.Left
        End If
        If PrevNewW <> NewWidth Then
            SendMessageA hWnd, LVM_SETCOLUMNWIDTH, CLng(DraggingIndex), NewWidth
            Call UpdateHeaders(True)
            PrevNewW = NewWidth
        End If
    End If
End Sub

Private Sub tmrUpdateScrollBars_Timer()
    On Error Resume Next
    Dim rtn     As Long
    Dim lstRect As RECT
    
    If Not Ambient.UserMode Then
        UserControl.tmrUpdateScrollBars.Enabled = False
    End If
    
    GetWindowRect hWnd, lstRect
    rtn = SendMessageA(hWnd, LVM_APPROXIMATEVIEWRECT, -1, 0)
    ReqW = LoWord(rtn)
    ReqH = HiWord(rtn)
    
    If ReqW > lstRect.Right - lstRect.Left Then
        If UserControl.HScrollBar.Visible = False Then
            UserControl.HScrollBar.Visible = True
            Call UserControl_Resize
        End If
        If UserControl.HScrollBar.Max <> ReqW - lstRect.Right + lstRect.Left Then
            UserControl.HScrollBar.Max = ReqW - lstRect.Right + lstRect.Left
        End If
        If UserControl.HScrollBar.Value > UserControl.HScrollBar.Max Then
            UserControl.HScrollBar.Value = UserControl.HScrollBar.Max
        End If
        If UserControl.HScrollBar.BarWidth <> CLng((UserControl.HScrollBar.Width - 480 * 2) / ReqW * (lstRect.Right - lstRect.Left)) Then
            UserControl.HScrollBar.BarWidth = CLng((UserControl.HScrollBar.Width - 480 * 2) / ReqW * (lstRect.Right - lstRect.Left))
        End If
    Else
        If UserControl.HScrollBar.Visible = True Then
            UserControl.HScrollBar.Visible = False
            UserControl.HScrollBar.Value = 0
            Call UserControl_Resize
        End If
    End If
    If ReqH > lstRect.bottom - lstRect.Top Then
        Dim ItemRect    As RECT
        Dim TopIndex    As Long
        
        TopIndex = SendMessageA(hWnd, LVM_GETTOPINDEX, 0, 0)
        SendMessageA hWnd, LVM_GETITEMRECT, TopIndex, ByVal VarPtr(ItemRect)
        
        If UserControl.VscrollBar.Visible = False Then
            UserControl.VscrollBar.Visible = True
            Call UserControl_Resize
        End If
        If UserControl.VscrollBar.Max <> ReqH - lstRect.bottom + lstRect.Top Then
            UserControl.VscrollBar.Max = ReqH - lstRect.bottom + lstRect.Top
            UserControl.VscrollBar.SmallChange = (ItemRect.bottom - ItemRect.Top)
            UserControl.VscrollBar.LargeChange = UserControl.VscrollBar.SmallChange * 3
        End If
        If UserControl.VscrollBar.BarHeight <> CLng((UserControl.VscrollBar.Height - 480 * 2) / ReqH * (lstRect.bottom - lstRect.Top)) Then
            UserControl.VscrollBar.BarHeight = CLng((UserControl.VscrollBar.Height - 480 * 2) / ReqH * (lstRect.bottom - lstRect.Top))
        End If
        If Abs(UserControl.VscrollBar.Value - (ItemRect.bottom - ItemRect.Top) * TopIndex) > (ItemRect.bottom - ItemRect.Top) Then
            UserControl.VscrollBar.Value = (ItemRect.bottom - ItemRect.Top) * TopIndex
        End If
    Else
        If UserControl.VscrollBar.Visible = True Then
            UserControl.VscrollBar.Visible = False
            UserControl.VscrollBar.Value = 0
            Call UserControl_Resize
        End If
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim li  As LVITEM
    
    If KeyCode = vbKeyUp Then
        li.iItem = SendMessageA(hWnd, LVM_GETNEXTITEM, -1, LVNI_SELECTED) - 1
        li.state = LVIS_FOCUSED Or LVIS_SELECTED
        li.mask = LVIF_STATE
        li.stateMask = LVIS_FOCUSED Or LVIS_SELECTED
        SendMessageA hWnd, LVM_SETITEMSTATE, li.iItem, ByVal VarPtr(li)
    ElseIf KeyCode = vbKeyDown Then
        li.iItem = SendMessageA(hWnd, LVM_GETNEXTITEM, -1, LVNI_SELECTED) + 1
        li.state = LVIS_FOCUSED Or LVIS_SELECTED
        li.mask = LVIF_STATE
        li.stateMask = LVIS_FOCUSED Or LVIS_SELECTED
        SendMessageA hWnd, LVM_SETITEMSTATE, li.iItem, ByVal VarPtr(li)
    End If
    EnsureVisible li.iItem, True
    EnsureVisible li.iItem, False
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub tmrCheckFocusAndColumn_Timer()
    If NeedToCheckFocus Then
        Dim cur     As POINT
        Dim Target  As Long
        
        GetCursorPos cur
        Target = WindowFromPoint(cur.X, cur.Y)
        If Target <> UserControl.hWnd Then
            Dim i   As Integer
            
            For i = 0 To UserControl.labColumnHeader.UBound
                UserControl.labColumnHeader(i).BackColor = RGB(62, 62, 66)
            Next i
            NeedToCheckFocus = False
            PrevIndex = -1
        End If
    End If
    
    Call UpdateHeaders
End Sub
 
Private Sub UserControl_Initialize()
    hWnd = CreateWindowExA(WS_EX_NOPARENTNOTIFY, "SysListView32", "", _
        WS_VISIBLE Or WS_CHILD Or WS_BORDER Or WS_TABSTOP Or LVS_ALIGNLEFT Or LVS_REPORT Or LVS_SINGLESEL Or LVS_NOCOLUMNHEADER, _
        -1, labColumnHeader(0).Height / Screen.TwipsPerPixelY, UserControl.Width / Screen.TwipsPerPixelX, _
        UserControl.Height / Screen.TwipsPerPixelY, UserControl.hWnd, 0, App.hInstance, 0)
    
    SetPropA hWnd, "ID", CtlListPushBack(Me)
    SetPropA hWnd, "PARENT_CTL", UserControl.hWnd
    
    SendMessageA hWnd, LVM_SETBKCOLOR, ByVal 0, ByVal RGB(51, 51, 55)
    SendMessageA hWnd, LVM_SETTEXTBKCOLOR, ByVal 0, ByVal RGB(51, 51, 55)
    SendMessageA hWnd, LVM_SETTEXTCOLOR, ByVal 0, ByVal RGB(240, 240, 240)
    
    PrevIndex = -1
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If NeedToCheckFocus Then
        Dim i   As Integer
        
        For i = 0 To UserControl.labColumnHeader.UBound
            UserControl.labColumnHeader(i).BackColor = RGB(62, 62, 66)
        Next i
        NeedToCheckFocus = False
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    If Ambient.UserMode Then
        PrevLVUserCtlProc = SetWindowLongA(UserControl.hWnd, GWL_WNDPROC, AddressOf ListViewNotifyMessageProc)
        PrevListViewProc = SetWindowLongA(hWnd, GWL_WNDPROC, AddressOf ListViewProc)
        UserControl.tmrCheckFocusAndColumn.Enabled = True
    End If
    m_FullRowSelect = PropBag.ReadProperty("FullRowSelect", m_def_FullRowSelect)
    If m_FullRowSelect Then
        SendMessageA hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, LVS_EX_FULLROWSELECT
    End If
End Sub

Private Sub UserControl_Resize()
    SetWindowPos hWnd, 0, -1, (labColumnHeader(0).Height - 15) / Screen.TwipsPerPixelY - 1, _
        IIf(UserControl.VscrollBar.Visible, (UserControl.Width - UserControl.VscrollBar.Width) / Screen.TwipsPerPixelX, UserControl.Width / Screen.TwipsPerPixelX), _
        IIf(UserControl.HScrollBar.Visible, (UserControl.Height - UserControl.labColumnHeader(0).Height - UserControl.HScrollBar.Height) / Screen.TwipsPerPixelY + 1, _
            (UserControl.Height - UserControl.labColumnHeader(0).Height) / Screen.TwipsPerPixelY + 1), SWP_NOZORDER
    
    If UserControl.VscrollBar.Visible Then
        UserControl.VscrollBar.Top = UserControl.labColumnHeader(0).Height
        UserControl.VscrollBar.Left = UserControl.Width - UserControl.VscrollBar.Width
        UserControl.VscrollBar.Height = IIf(UserControl.HScrollBar.Visible, _
            UserControl.Height - UserControl.HScrollBar.Height - UserControl.labColumnHeader(0).Height, _
            UserControl.Height - UserControl.labColumnHeader(0).Height)
    End If
    If UserControl.HScrollBar.Visible Then
        UserControl.HScrollBar.Left = 0
        UserControl.HScrollBar.Top = UserControl.Height - UserControl.HScrollBar.Height
        UserControl.HScrollBar.Width = IIf(UserControl.VscrollBar.Visible, UserControl.Width - UserControl.VscrollBar.Width, UserControl.Width)
    End If
    
    Dim i   As Integer
    For i = 0 To UserControl.labColumnHeader.UBound
        If i = 0 Then
            UserControl.labColumnHeader(0).Left = -15 - UserControl.HScrollBar.Value * Screen.TwipsPerPixelX
            UserControl.labColumnHeader(0).Top = -15
            UserControl.picResizer(0).Top = -15
            UserControl.picResizer(0).Left = UserControl.labColumnHeader(0).Left + UserControl.labColumnHeader(0).Width - 60
            UserControl.picResizer(0).Width = 90
            UserControl.picResizer(0).Height = UserControl.labColumnHeader(0).Height
        Else
            UserControl.labColumnHeader(i).Left = UserControl.labColumnHeader(i - 1).Left + UserControl.labColumnHeader(i - 1).Width + 15
            UserControl.labColumnHeader(i).Top = -15
            UserControl.picResizer(i).Top = -15
            UserControl.picResizer(i).Left = UserControl.labColumnHeader(i).Left + UserControl.labColumnHeader(i).Width - 60
            UserControl.picResizer(i).Width = 90
            UserControl.picResizer(i).Height = UserControl.labColumnHeader(0).Height
        End If
    Next i
End Sub

Private Sub VScrollBar_ValueChanged(NewValue As Long)
    LockWindowUpdate hWnd
    SendMessageA hWnd, LVM_SCROLL, 0, -ReqH
    SendMessageA hWnd, LVM_SCROLL, 0, NewValue
    LockWindowUpdate 0
End Sub

Private Sub HScrollBar_ValueChanged(NewValue As Long)
    LockWindowUpdate hWnd
    SendMessageA hWnd, LVM_SCROLL, -ReqW, 0
    SendMessageA hWnd, LVM_SCROLL, NewValue, 0
    Call UserControl_Resize
    LockWindowUpdate 0
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
        SendMessageA hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, LVS_EX_FULLROWSELECT
    Else
        SendMessageA hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, 0
    End If
    PropertyChanged "FullRowSelect"
End Property

Public Property Get ListViewHwnd() As Long
    ListViewHwnd = hWnd
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FullRowSelect = m_def_FullRowSelect
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("FullRowSelect", m_FullRowSelect, m_def_FullRowSelect)
End Sub


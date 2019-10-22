VERSION 5.00
Begin VB.UserControl DarkTreeView 
   BackColor       =   &H00302D2D&
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   ScaleHeight     =   3960
   ScaleWidth      =   4065
   ToolboxBitmap   =   "DarkTreeView.ctx":0000
End
Attribute VB_Name = "DarkTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================================
'描述:      树视图控件。因为VB6的树视图ActiveX控件有Bug，只好自己写一个了
'作者:      冰棍
'文件:      DarkTreeView.ctl
'====================================================

Option Explicit

'下面这些事件自己写WndProc来调用吧... 我也很绝望啊... o(╥﹏╥)o
'因为我实在是没办法用优雅的方式从模块触发这些事件... 反正树视图也只用一次，干脆就直接在模块里写事件触发算了
Event Click(ByRef bCancel As Boolean)
Event RightClick(ByRef bCancel As Boolean)
Event BeginLabelEdit(ByVal hTreeItem As Long, ByRef bCancel As Boolean)
Event EndLabelEdit(ByVal hTreeItem As Long, NewText As String, ByRef bCancel As Boolean)
Event ItemExpanding(ByVal hTreeItem As Long, ByRef bCancel As Boolean)
Event KeyDown(ByVal KeyCode As Long, ByVal IsLongPress As Boolean)
Event KeyUp(ByVal KeyCode As Long)
Event MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
Event MouseDown(ByVal Button As Long, ByVal X As Long, ByVal Y As Long)
Event MouseUp(ByVal Button As Long, ByVal X As Long, ByVal Y As Long)
Event DoubleClick(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
Event SelChanging(ByVal hOldTreeItem As Long, ByVal hNewTreeItem As Long, ByRef bCancel As Boolean)
Event SelChanged(ByVal hOldTreeItem As Long, ByVal hNewTreeItem As Long)

Dim wndTreeView     As Long                                                         '树视图控件的hWnd

Private Sub UserControl_Initialize()
    '创建树视图控件
    wndTreeView = CreateWindowExA(0, "SysTreeView32", "", _
        WS_VISIBLE Or WS_CHILD Or TVS_HASBUTTONS Or TVS_SHOWSELALWAYS Or TVS_EDITLABELS Or TVS_FULLROWSELECT Or TVS_HASLINES Or TVS_LINESATROOT Or TVS_LINESATROOT, _
        0, 0, 100, 300, UserControl.hWnd, 0, App.hInstance, 0)  'Or TVS_HASLINES

    '设置控件颜色
    SendMessageA wndTreeView, TVM_SETBKCOLOR, 0, ByVal &H302D2D
    SendMessageA wndTreeView, TVM_SETTEXTCOLOR, 0, ByVal &HF0F0F0
    SendMessageA wndTreeView, TVM_SETLINECOLOR, 0, ByVal &H808080
End Sub

Private Sub UserControl_Resize()
    MoveWindow wndTreeView, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY, 1
End Sub

'描述:      添加项目到树视图中
'参数:      ItemText: 项目的文本
'.          ParentItem: 父节点的句柄
'返回值:    新创建的项目的句柄
Public Function AddItem(ItemText As String, Optional ParentItem As Long = 0) As Long
    Dim ti          As TVINSERTSTRUCTEX
    Dim TextBuf()   As Byte
    
    TextBuf = StrConvEx(ItemText)
    With ti
        .hInsertAfter = TVI_LAST
        .hParent = ParentItem
        .itemex.mask = TVIF_TEXT
        .itemex.pszText = VarPtr(TextBuf(0))
        .itemex.cchTextMax = UBound(TextBuf)
    End With
    AddItem = SendMessageA(wndTreeView, TVM_INSERTITEM, 0, ByVal VarPtr(ti))
    
    '设置控件颜色
    SendMessageA wndTreeView, TVM_SETBKCOLOR, 0, ByVal &H302D2D
    SendMessageA wndTreeView, TVM_SETTEXTCOLOR, 0, ByVal &HF0F0F0
    SendMessageA wndTreeView, TVM_SETLINECOLOR, 0, ByVal &H808080
End Function

'描述:      删除指定的项目
'参数:      Item: 需要删除的项目的句柄。设置为0则删除所有的项目
'返回值:    若删除成功则返回非0的整数，否则返回0
Public Function RemoveItem(ByVal Item As Long) As Boolean
    RemoveItem = (SendMessageA(wndTreeView, TVM_DELETEITEM, 0, ByVal Item) <> 0)
End Function

'描述:      确保指定的项目可视
'参数:      Item: 指定项目句柄
Public Sub EnsureVisible(ByVal Item As Long)
    SendMessageA wndTreeView, TVM_ENSUREVISIBLE, 0, ByVal Item
End Sub

'描述:      展开或者收缩树状图
'参数:      Item: 需要被展开或者收缩的列表项
'.          Mode: 展开或者收缩。1: 收缩; 2: 展开; 3: 切换展开或者收缩
'返回值:    如果成功，返回True
Public Function ExpandItems(ByVal Item As Long, Mode As Integer) As Boolean
    ExpandItems = (SendMessageA(wndTreeView, TVM_EXPAND, ByVal Mode, ByVal Item) <> 0)
End Function

'描述:      开始编辑文本
'参数:      Item: 需要编辑文本的列表项
Public Function EditLabel(ByVal Item As Long) As Boolean
    EditLabel = (SendMessageA(wndTreeView, TVM_EDITLABEL, 0, ByVal Item) <> 0)
End Function

'描述:      取消编辑文本
'参数:      SaveChanges: 是否保存对项目的修改
'返回值:    若执行成功则返回True
Public Function EndEditLabel(SaveChanges As Boolean) As Boolean
    EndEditLabel = (SendMessageA(wndTreeView, TVM_ENDEDITLABELNOW, CLng(SaveChanges), 0) <> 0)
End Function

'描述:      获取指定列表项的文本
'参数:      Item: 列表项的句柄
'返回值:    指定列表项的文本
Public Function GetItemText(ByVal Item As Long) As String
    Dim tmp(260)    As Byte
    Dim tvi         As TVITEM
    
    With tvi
        .mask = TVIF_TEXT
        .cchTextMax = 260
        .pszText = VarPtr(tmp(0))
        .hItem = Item
    End With
    SendMessageA wndTreeView, TVM_GETITEM, 0, ByVal VarPtr(tvi)
    GetItemText = ByteArrayConv(tmp)
End Function

'描述:      获取指定列表项的文本
'参数:      Item: 列表项的句柄
'.          NewText: 新的文本
'返回值:    指定列表项的文本
Public Function SetItemText(ByVal Item As Long, NewText As String) As Boolean
    Dim tvi         As TVITEM
    Dim buf()       As Byte
    
    buf = StrConvEx(NewText)
    With tvi
        .mask = TVIF_TEXT
        .cchTextMax = UBound(buf)
        .pszText = VarPtr(buf(0))
        .hItem = Item
    End With
    SetItemText = (SendMessageA(wndTreeView, TVM_SETITEM, 0, ByVal VarPtr(tvi)) <> 0)
End Function

'描述:      获取选择的项目句柄
'返回值:    选择的项目句柄。如果没有选择项目则返回0
Public Function GetSelectedItem() As Long
    GetSelectedItem = SendMessageA(wndTreeView, TVM_GETNEXTITEM, TVGN_CARET, 0)
End Function

'描述:      获取指定列表项的根节点句柄
'参数:      Item: 列表项的句柄
'返回值:    指定列表项的根节点句柄。若没有选择项目或者选择的项目无效，则返回0
Public Function GetParentItem(ByVal Item As Long) As Long
    GetParentItem = SendMessageA(wndTreeView, TVM_GETNEXTITEM, TVGN_PARENT, ByVal Item)
End Function

'描述:      选择指定的列表项目
'参数:      Item: 列表项的句柄
'返回值:    如果执行成功则返回True
Public Function SelectItem(ByVal Item As Long) As Boolean
    SelectItem = (SendMessageA(wndTreeView, TVM_SELECTITEM, TVGN_CARET, ByVal Item) <> 0)
End Function

'描述:      从指定坐标获取列表项的句柄
'参数:      X, Y: 指定坐标
'返回值:    如果有列表项在指定的坐标的位置，返回该列表项的句柄；否则返回0
Public Function HitTest(X As Long, Y As Long) As Long
    Dim tvhti   As TVHITTESTINFO
    
    tvhti.pt.X = X
    tvhti.pt.Y = Y
    HitTest = SendMessageA(wndTreeView, TVM_HITTEST, ByVal 0, ByVal VarPtr(tvhti))
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    '设置控件子类化
    If Ambient.UserMode Then
        SetPropA wndTreeView, "PrevWndProc", SetWindowLongA(wndTreeView, GWL_WNDPROC, AddressOf TreeViewWindowProc)
        SetPropA UserControl.hWnd, "PrevWndProc", SetWindowLongA(UserControl.hWnd, GWL_WNDPROC, AddressOf TreeViewUserCtlWindowProc)
    End If
    
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get TreeViewHwnd() As Long
    TreeViewHwnd = wndTreeView
End Property


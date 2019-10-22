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
'����:      ����ͼ�ؼ�����ΪVB6������ͼActiveX�ؼ���Bug��ֻ���Լ�дһ����
'����:      ����
'�ļ�:      DarkTreeView.ctl
'====================================================

Option Explicit

'������Щ�¼��Լ�дWndProc�����ð�... ��Ҳ�ܾ�����... o(�i�n�i)o
'��Ϊ��ʵ����û�취�����ŵķ�ʽ��ģ�鴥����Щ�¼�... ��������ͼҲֻ��һ�Σ��ɴ��ֱ����ģ����д�¼���������
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

Dim wndTreeView     As Long                                                         '����ͼ�ؼ���hWnd

Private Sub UserControl_Initialize()
    '��������ͼ�ؼ�
    wndTreeView = CreateWindowExA(0, "SysTreeView32", "", _
        WS_VISIBLE Or WS_CHILD Or TVS_HASBUTTONS Or TVS_SHOWSELALWAYS Or TVS_EDITLABELS Or TVS_FULLROWSELECT Or TVS_HASLINES Or TVS_LINESATROOT Or TVS_LINESATROOT, _
        0, 0, 100, 300, UserControl.hWnd, 0, App.hInstance, 0)  'Or TVS_HASLINES

    '���ÿؼ���ɫ
    SendMessageA wndTreeView, TVM_SETBKCOLOR, 0, ByVal &H302D2D
    SendMessageA wndTreeView, TVM_SETTEXTCOLOR, 0, ByVal &HF0F0F0
    SendMessageA wndTreeView, TVM_SETLINECOLOR, 0, ByVal &H808080
End Sub

Private Sub UserControl_Resize()
    MoveWindow wndTreeView, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY, 1
End Sub

'����:      �����Ŀ������ͼ��
'����:      ItemText: ��Ŀ���ı�
'.          ParentItem: ���ڵ�ľ��
'����ֵ:    �´�������Ŀ�ľ��
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
    
    '���ÿؼ���ɫ
    SendMessageA wndTreeView, TVM_SETBKCOLOR, 0, ByVal &H302D2D
    SendMessageA wndTreeView, TVM_SETTEXTCOLOR, 0, ByVal &HF0F0F0
    SendMessageA wndTreeView, TVM_SETLINECOLOR, 0, ByVal &H808080
End Function

'����:      ɾ��ָ������Ŀ
'����:      Item: ��Ҫɾ������Ŀ�ľ��������Ϊ0��ɾ�����е���Ŀ
'����ֵ:    ��ɾ���ɹ��򷵻ط�0�����������򷵻�0
Public Function RemoveItem(ByVal Item As Long) As Boolean
    RemoveItem = (SendMessageA(wndTreeView, TVM_DELETEITEM, 0, ByVal Item) <> 0)
End Function

'����:      ȷ��ָ������Ŀ����
'����:      Item: ָ����Ŀ���
Public Sub EnsureVisible(ByVal Item As Long)
    SendMessageA wndTreeView, TVM_ENSUREVISIBLE, 0, ByVal Item
End Sub

'����:      չ������������״ͼ
'����:      Item: ��Ҫ��չ�������������б���
'.          Mode: չ������������1: ����; 2: չ��; 3: �л�չ����������
'����ֵ:    ����ɹ�������True
Public Function ExpandItems(ByVal Item As Long, Mode As Integer) As Boolean
    ExpandItems = (SendMessageA(wndTreeView, TVM_EXPAND, ByVal Mode, ByVal Item) <> 0)
End Function

'����:      ��ʼ�༭�ı�
'����:      Item: ��Ҫ�༭�ı����б���
Public Function EditLabel(ByVal Item As Long) As Boolean
    EditLabel = (SendMessageA(wndTreeView, TVM_EDITLABEL, 0, ByVal Item) <> 0)
End Function

'����:      ȡ���༭�ı�
'����:      SaveChanges: �Ƿ񱣴����Ŀ���޸�
'����ֵ:    ��ִ�гɹ��򷵻�True
Public Function EndEditLabel(SaveChanges As Boolean) As Boolean
    EndEditLabel = (SendMessageA(wndTreeView, TVM_ENDEDITLABELNOW, CLng(SaveChanges), 0) <> 0)
End Function

'����:      ��ȡָ���б�����ı�
'����:      Item: �б���ľ��
'����ֵ:    ָ���б�����ı�
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

'����:      ��ȡָ���б�����ı�
'����:      Item: �б���ľ��
'.          NewText: �µ��ı�
'����ֵ:    ָ���б�����ı�
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

'����:      ��ȡѡ�����Ŀ���
'����ֵ:    ѡ�����Ŀ��������û��ѡ����Ŀ�򷵻�0
Public Function GetSelectedItem() As Long
    GetSelectedItem = SendMessageA(wndTreeView, TVM_GETNEXTITEM, TVGN_CARET, 0)
End Function

'����:      ��ȡָ���б���ĸ��ڵ���
'����:      Item: �б���ľ��
'����ֵ:    ָ���б���ĸ��ڵ�������û��ѡ����Ŀ����ѡ�����Ŀ��Ч���򷵻�0
Public Function GetParentItem(ByVal Item As Long) As Long
    GetParentItem = SendMessageA(wndTreeView, TVM_GETNEXTITEM, TVGN_PARENT, ByVal Item)
End Function

'����:      ѡ��ָ�����б���Ŀ
'����:      Item: �б���ľ��
'����ֵ:    ���ִ�гɹ��򷵻�True
Public Function SelectItem(ByVal Item As Long) As Boolean
    SelectItem = (SendMessageA(wndTreeView, TVM_SELECTITEM, TVGN_CARET, ByVal Item) <> 0)
End Function

'����:      ��ָ�������ȡ�б���ľ��
'����:      X, Y: ָ������
'����ֵ:    ������б�����ָ���������λ�ã����ظ��б���ľ�������򷵻�0
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
    
    '���ÿؼ����໯
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


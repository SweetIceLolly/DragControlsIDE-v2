VERSION 5.00
Begin VB.UserControl DarkListBox 
   Appearance      =   0  'Flat
   BackColor       =   &H00333333&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1890
   ScaleHeight     =   2550
   ScaleWidth      =   1890
   ToolboxBitmap   =   "DarkListBox.ctx":0000
   Begin VB.Timer tmrUpdateVScrollBar 
      Interval        =   100
      Left            =   600
      Top             =   1920
   End
   Begin DragControlsIDE.DarkVScrollBar VScrollBar 
      Height          =   2295
      Left            =   1320
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4048
   End
   Begin VB.ListBox lstMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00373333&
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   1710
      ItemData        =   "DarkListBox.ctx":0312
      Left            =   0
      List            =   "DarkListBox.ctx":0314
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "DarkListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark¡áListBox by IceLolly
'Date: 2018.8.26

'Event Declarations:
Event Click() 'MappingInfo=lstMain,lstMain,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=lstMain,lstMain,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=lstMain,lstMain,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=lstMain,lstMain,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=lstMain,lstMain,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lstMain,lstMain,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lstMain,lstMain,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lstMain,lstMain,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Scroll() 'MappingInfo=lstMain,lstMain,-1,Scroll
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."

Private Sub lstMain_Scroll()
    UserControl.VscrollBar.Value = UserControl.lstMain.TopIndex
    RaiseEvent Scroll
End Sub

Private Sub tmrUpdateVScrollBar_Timer()
    On Error Resume Next
    Dim ItemsPerPage    As Integer
    Dim ItemCount       As Integer
    
    If Not Ambient.UserMode Then
        UserControl.tmrUpdateVScrollBar.Enabled = False
    End If
    
    ItemsPerPage = UserControl.lstMain.Height \ _
        (SendMessageA(UserControl.lstMain.hWnd, LB_GETITEMHEIGHT, 0, 0) * Screen.TwipsPerPixelY)
    ItemCount = UserControl.lstMain.ListCount
    If ItemCount > ItemsPerPage Then
        If UserControl.VscrollBar.Max <> ItemCount - ItemsPerPage Then
            UserControl.VscrollBar.Max = ItemCount - ItemsPerPage
        End If
        If UserControl.VscrollBar.BarHeight > 0 And UserControl.VscrollBar.BarHeight < 120 Then
            UserControl.VscrollBar.BarHeight = 120
        ElseIf UserControl.VscrollBar.BarHeight = 120 Then
            Exit Sub
        ElseIf UserControl.VscrollBar.BarHeight <> CLng((UserControl.VscrollBar.Height - 480 * 2) / ItemCount * ItemsPerPage) Then
            UserControl.VscrollBar.BarHeight = (UserControl.VscrollBar.Height - 480 * 2) / ItemCount * ItemsPerPage
        End If
        UserControl.VscrollBar.Enabled = True
    Else
        If UserControl.VscrollBar.BarHeight <> 0 Then
            UserControl.VscrollBar.BarHeight = 0
        End If
        UserControl.VscrollBar.Enabled = False
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lstMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lstMain.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    lstMain.Text = PropBag.ReadProperty("Text", "")
    lstMain.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    lstMain.TopIndex = PropBag.ReadProperty("TopIndex", 0)
    UserControl.BorderStyle = IIf(PropBag.ReadProperty("HasBorder", True), 1, 0)
    
    SetWindowLongA UserControl.lstMain.hWnd, GWL_STYLE, GetWindowLongA(UserControl.lstMain.hWnd, GWL_STYLE) And Not WS_BORDER
    If Ambient.UserMode Then
        PrevUserCtlProc = SetWindowLongA(UserControl.hWnd, GWL_WNDPROC, AddressOf ListBoxRedrawProc)
        PrevListBoxProc = SetWindowLongA(UserControl.lstMain.hWnd, GWL_WNDPROC, AddressOf ListBoxWheelFixProc)
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    UserControl.lstMain.Height = UserControl.Height
    UserControl.lstMain.Width = UserControl.Width - UserControl.VscrollBar.Width
    UserControl.VscrollBar.Left = UserControl.lstMain.Width
    UserControl.Height = UserControl.lstMain.Height
    UserControl.VscrollBar.Height = UserControl.Height
    UserControl.Width = UserControl.VscrollBar.Left + UserControl.VscrollBar.Width
End Sub

Private Sub VScrollBar_ValueChanged(NewValue As Long)
    UserControl.lstMain.TopIndex = NewValue
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,AddItem
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    lstMain.AddItem Item, Index
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Clear
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of a control or the system Clipboard."
    lstMain.Clear
End Sub

Private Sub lstMain_Click()
    RaiseEvent Click
End Sub

Private Sub lstMain_DblClick()
    RaiseEvent DblClick
End Sub

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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lstMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lstMain.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = lstMain.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,ListCount
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = lstMain.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = lstMain.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    lstMain.ListIndex = New_ListIndex
    PropertyChanged "ListIndex"
End Property

Private Sub lstMain_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub lstMain_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub lstMain_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub lstMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lstMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lstMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,NewIndex
Public Property Get NewIndex() As Integer
Attribute NewIndex.VB_Description = "Returns the index of the item most recently added to a control."
    NewIndex = lstMain.NewIndex
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,RemoveItem
Public Sub RemoveItem(ByVal Index As Integer)
Attribute RemoveItem.VB_Description = "Removes an item from a ListBox or ComboBox control or a row from a Grid control."
    lstMain.RemoveItem Index
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    lstMain.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Sorted
Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = lstMain.Sorted
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = lstMain.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    lstMain.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = lstMain.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    lstMain.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,TopIndex
Public Property Get TopIndex() As Integer
Attribute TopIndex.VB_Description = "Returns/sets which item in a control is displayed in the topmost position."
    TopIndex = lstMain.TopIndex
End Property

Public Property Let TopIndex(ByVal New_TopIndex As Integer)
    lstMain.TopIndex() = New_TopIndex
    PropertyChanged "TopIndex"
End Property

Public Property Get List(ByVal Index As Integer) As String
    List = UserControl.lstMain.List(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal NewText As String)
    UserControl.lstMain.List(Index) = NewText
End Property

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lstMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("ListIndex", lstMain.ListIndex, 0)
    Call PropBag.WriteProperty("Text", lstMain.Text, "")
    Call PropBag.WriteProperty("ToolTipText", lstMain.ToolTipText, "")
    Call PropBag.WriteProperty("TopIndex", lstMain.TopIndex, 0)
    Call PropBag.WriteProperty("HasBorder", IIf(UserControl.BorderStyle = 1, True, False), True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get HasBorder() As Boolean
Attribute HasBorder.VB_Description = "Returns/sets the border style for an object."
    HasBorder = IIf(UserControl.BorderStyle = 1, True, False)
End Property

Public Property Let HasBorder(ByVal New_HasBorder As Boolean)
    UserControl.BorderStyle() = IIf(New_HasBorder, 1, 0)
    PropertyChanged "HasBorder"
End Property


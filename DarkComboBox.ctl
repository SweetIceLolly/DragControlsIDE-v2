VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Image.ocx"
Begin VB.UserControl DarkComboBox 
   BackColor       =   &H00463F3F&
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   PropertyPages   =   "DarkComboBox.ctx":0000
   ScaleHeight     =   930
   ScaleWidth      =   3015
   ToolboxBitmap   =   "DarkComboBox.ctx":0013
   Begin VB.Timer tmrCheckFocus 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin DragControlsIDE.DarkEdit edMain 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Dark°·ComboBox"
   End
   Begin ImageX.aicAlphaImage imgDropDown 
      Height          =   360
      Left            =   1680
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Image           =   "DarkComboBox.ctx":0325
      Enabled         =   0   'False
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00CC7A00&
      Height          =   375
      Left            =   1680
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgMouseDown 
      Height          =   360
      Left            =   1560
      Picture         =   "DarkComboBox.ctx":03D3
      Top             =   480
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgMouseIn 
      Height          =   360
      Left            =   1080
      Picture         =   "DarkComboBox.ctx":0B3D
      Top             =   480
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgNormal 
      Height          =   360
      Left            =   600
      Picture         =   "DarkComboBox.ctx":12A7
      Top             =   480
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "DarkComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'Dark°·ComboBox by IceLolly
'Date: 2018.8.10

'Back           R    G    B
'Normal         63,  63,  70
'Mouse in       45,  45,  48
'Mouse down     0,   122, 204

Dim bDown           As Boolean
Dim bDropDown       As Boolean
Dim ListItems()     As String               'Base = 1
Dim CurrListIndex   As Integer              'Base = 0

'Default Property Values:
Const m_def_Editable = True
Const m_def_ListHeight = 2000
'Property Variables:
Dim m_ListCount As Integer
Dim m_Editable As Boolean
Dim m_ListHeight As Variant
'Event Declarations:
Event Changed(NewText As String, CurrIndex As Integer)

Public Sub ShowList()
    Call UserControl_MouseDown(1, 0, 0, 0)
End Sub

Public Sub HideList()
    bDropDown = False
    Unload frmComboBoxListWindow
    UserControl.imgDropDown.LoadImage_FromStdPicture UserControl.imgNormal.Picture
    UserControl.BackColor = RGB(63, 63, 70)
    UserControl.shpBorder.Visible = False
End Sub

Public Sub Clear()
    ReDim ListItems(0)
    Items(0) = ""
End Sub

Public Sub RemoveItem(ByVal ItemRemove As Integer)
    Dim i       As Integer
    
    For i = ItemRemove To UBound(ListItems) - 1
        ListItems(i) = ListItems(i + 1)
    Next i
    ReDim Preserve ListItems(UBound(ListItems) - 1)
End Sub

Public Sub AddItem(ByVal NewItem As String)
    ReDim Preserve ListItems(UBound(ListItems) + 1)
    Items(UBound(ListItems)) = NewItem
End Sub

Private Sub edMain_Change()
    RaiseEvent Changed(UserControl.edMain.Text, CurrListIndex)
End Sub

Private Sub edMain_GotFocus()
    UserControl.edMain.SelStart = 0
    UserControl.edMain.SelLength = Len(UserControl.edMain.Text)
End Sub

Private Sub edMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        Me.ListIndex = Me.ListIndex + 1
    ElseIf KeyCode = vbKeyUp Then
        Me.ListIndex = Me.ListIndex - 1
    End If
End Sub

Private Sub edMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Editable Then
        Me.ShowList
    End If
End Sub

Private Sub tmrCheckFocus_Timer()
    Dim pt      As POINT
    Dim Target  As Long
    
    If bDropDown Then
        Exit Sub
    End If
    GetCursorPos pt
    Target = WindowFromPoint(pt.X, pt.Y)
    If Target <> UserControl.hWnd Then
        UserControl.imgDropDown.LoadImage_FromStdPicture UserControl.imgNormal.Picture
        UserControl.BackColor = RGB(63, 63, 70)
        UserControl.shpBorder.Visible = False
    End If
End Sub

Private Sub UserControl_InitProperties()
    ReDim ListItems(0)
    m_Editable = m_def_Editable
    m_ListHeight = m_def_ListHeight
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    UserControl.imgDropDown.Left = UserControl.Width - UserControl.imgDropDown.Width
    UserControl.edMain.Width = UserControl.Width - UserControl.imgDropDown.Width
    UserControl.edMain.Height = UserControl.Height
    UserControl.imgDropDown.Top = UserControl.Height / 2 - UserControl.imgDropDown.Height / 2
    UserControl.shpBorder.Left = UserControl.imgDropDown.Left
    UserControl.shpBorder.Width = UserControl.imgDropDown.Width
    UserControl.shpBorder.Height = UserControl.Height
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        bDown = True
        bDropDown = Not bDropDown
        UserControl.imgDropDown.LoadImage_FromStdPicture UserControl.imgMouseDown.Picture
        UserControl.BackColor = RGB(0, 122, 204)
        UserControl.shpBorder.Visible = False
        
        '---------------------------------------------
        If bDropDown Then
            Dim wRect   As RECT
            Dim i       As Integer
            
            GetWindowRect UserControl.hWnd, wRect
            Unload frmComboBoxListWindow
            With frmComboBoxListWindow
                Set .BoundCtl = Me
                .Left = wRect.Left * Screen.TwipsPerPixelX
                .Top = wRect.bottom * Screen.TwipsPerPixelY
                .MaxWidth = (wRect.Right - wRect.Left) * Screen.TwipsPerPixelX
                For i = 1 To UBound(ListItems)
                    .AddItem " " & ListItems(i)
                Next i
                .picContainer.Height = Me.ListHeight
                If .labItem(.labItem.UBound).Top + .labItem(.labItem.UBound).Height > .picContainer.Height Then
                    .picContainer.Height = .labItem(.labItem.UBound).Top + .labItem(.labItem.UBound).Height
                    .VscrollBar.Visible = True
                    .MaxWidth = .MaxWidth + .VscrollBar.Width
                    .VscrollBar.Left = .MaxWidth - .VscrollBar.Width
                    .VscrollBar.Max = .picContainer.Height - Me.ListHeight
                    .VscrollBar.Height = Me.ListHeight
                    .VscrollBar.BarHeight = (.VscrollBar.Height - 480 * 2) * Me.ListHeight / .picContainer.Height
                    If .VscrollBar.BarHeight < 120 Then
                        .VscrollBar.BarHeight = 120
                    End If
                    .VscrollBar.SmallChange = .labItem(0).Height
                    .VscrollBar.LargeChange = .labItem(0).Height * (Me.ListHeight \ .labItem(0).Height)
                    .picContainer.Width = .VscrollBar.Left
                    .Height = Me.ListHeight
                Else
                    .Height = .labItem(.labItem.UBound).Top + .labItem(.labItem.UBound).Height
                    .picContainer.Width = .MaxWidth
                End If
                .Width = .MaxWidth
                .Show
            End With
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bDown And Not bDropDown Then
        UserControl.imgDropDown.LoadImage_FromStdPicture UserControl.imgMouseIn.Picture
        UserControl.BackColor = RGB(45, 45, 48)
        UserControl.shpBorder.Visible = True
        UserControl.tmrCheckFocus.Enabled = True
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDown = False
End Sub

Public Property Get Items(ByVal Index As Integer) As String
    Items = ListItems(Index)
End Property

Public Property Let Items(ByVal Index As Integer, ByVal NewText As String)
    ListItems(Index) = NewText
    PropertyChanged "Items"
End Property

Public Property Get ListCount() As Integer
    ListCount = UBound(ListItems)
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim i       As Integer
    
    ReDim ListItems(PropBag.ReadProperty("ITEM_COUNT", 0))
    
    For i = 0 To UBound(ListItems)
        ListItems(i) = PropBag.ReadProperty("Items" & i, "")
    Next i
    
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    edMain.Text = PropBag.ReadProperty("Text", "Dark°·ComboBox")
    Set edMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Editable = PropBag.ReadProperty("Editable", m_def_Editable)
    m_ListHeight = PropBag.ReadProperty("ListHeight", m_def_ListHeight)
    
    If m_Editable Then
        UserControl.edMain.MousePointer = vbIbeam
        UserControl.edMain.Locked = False
    Else
        UserControl.edMain.MousePointer = vbArrow
        UserControl.edMain.Locked = True
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim i       As Integer
    
    For i = 0 To UBound(ListItems)
        PropBag.WriteProperty "Items" & i, ListItems(i)
    Next i
    PropBag.WriteProperty "ITEM_COUNT", ListCount
    
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Text", edMain.Text, "Dark°·ComboBox")
    Call PropBag.WriteProperty("Font", edMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("Editable", m_Editable, m_def_Editable)
    Call PropBag.WriteProperty("ListHeight", m_ListHeight, m_def_ListHeight)
End Sub

Public Property Get ListIndex() As Integer
    ListIndex = CurrListIndex
End Property

Public Property Let ListIndex(ByVal NewIndex As Integer)
    On Error Resume Next
    
    CurrListIndex = NewIndex
    If ListIndex > UBound(ListItems) Then
        UserControl.edMain.Text = ListItems(UBound(ListItems))
        CurrListIndex = UBound(ListItems)
    ElseIf ListIndex <= 0 Then
        If UBound(ListItems) > 0 Then
            UserControl.edMain.Text = ListItems(1)
        Else
            UserControl.edMain.Text = ""
        End If
        CurrListIndex = 1
    Else
        UserControl.edMain.Text = ListItems(NewIndex)
    End If
End Property

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
'MappingInfo=edMain,edMain,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = edMain.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    edMain.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=edMain,edMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = edMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set edMain.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Editable() As Boolean
Attribute Editable.VB_Description = "Returns/Sets if the user can edit the text."
    Editable = m_Editable
End Property

Public Property Let Editable(ByVal New_Editable As Boolean)
    m_Editable = New_Editable
    PropertyChanged "Editable"
    
    If New_Editable Then
        UserControl.edMain.MousePointer = vbIbeam
        UserControl.edMain.Locked = False
    Else
        UserControl.edMain.MousePointer = vbArrow
        UserControl.edMain.Locked = True
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ListHeight() As Long
Attribute ListHeight.VB_Description = "Returns/Sets the maximum height of the drop-down list."
    ListHeight = m_ListHeight
End Property

Public Property Let ListHeight(ByVal New_ListHeight As Long)
    m_ListHeight = New_ListHeight
    PropertyChanged "ListHeight"
End Property


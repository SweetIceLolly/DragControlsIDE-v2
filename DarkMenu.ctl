VERSION 5.00
Begin VB.UserControl DarkMenu 
   Alignable       =   -1  'True
   BackColor       =   &H00302D2D&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   PropertyPages   =   "DarkMenu.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   5565
   ToolboxBitmap   =   "DarkMenu.ctx":0014
   Begin VB.Timer tmrCheckFocus 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4200
      Top             =   240
   End
   Begin VB.Line lnBorderTop 
      BorderColor     =   &H00373333&
      Visible         =   0   'False
      X1              =   1920
      X2              =   1440
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line lnBorderLeft 
      BorderColor     =   &H00373333&
      Visible         =   0   'False
      X1              =   1440
      X2              =   1440
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line lnBorderRight 
      BorderColor     =   &H00373333&
      Visible         =   0   'False
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Label labRootItem 
      AutoSize        =   -1  'True
      BackColor       =   &H00302D2D&
      Caption         =   " Item"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9.75
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
      Width           =   480
   End
End
Attribute VB_Name = "DarkMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark°·Menu by IceLolly
'Date: 2018.8.13

'               R    G    B
'Normal         45,  45,  48
'Mouse in       62,  62,  64
'Mouse down     27,  27,  28

Private Type MenuItem
    MenuID          As Integer
    MenuText        As String
    SubMenus()      As String           'Base = 1
    SubMenuID()     As Integer          'Base = 1
    Enabled         As Boolean
    CheckBox        As Boolean
    Visible         As Boolean
    Checked         As Boolean          'Won't save in the property bag
    MenuIcon()      As Byte
End Type

Dim Menus()         As MenuItem         'Base = 1
Dim RootMenuID()    As Integer
Dim Levels()        As Integer
Dim bShow           As Boolean
Dim PrevShowMenu    As Integer

Dim PopupIndex      As Integer
Dim PrevX           As Single, _
    PrevY           As Single

'Default Property Values:
Const m_def_SpaceCount = 3
Const m_def_HeightAddition = 60
Const m_def_Transparent = True
'Property Variables:
Dim m_SpaceCount As Integer
Dim m_HeightAddition As Integer
Dim m_Transparent As Boolean
'Event Declarations:
Event MenuItemClicked(MenuID As Integer)

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Function MoveLeft()
    Dim i           As Label
    
    bShow = False
    For Each i In UserControl.labRootItem
        i.BackColor = RGB(45, 45, 48)
    Next i
    PopupIndex = PopupIndex - 1
    If PopupIndex < 0 Then
        PopupIndex = UserControl.labRootItem.UBound
    End If
    Call labRootItem_MouseDown(PopupIndex, vbLeftButton, 0, 0, 0)
    Call frmPopupMenu.Form_KeyDown(vbKeyDown, 0)
End Function

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Function MoveRight()
    Dim i           As Label
    
    bShow = False
    For Each i In UserControl.labRootItem
        i.BackColor = RGB(45, 45, 48)
    Next i
    PopupIndex = PopupIndex + 1
    If PopupIndex > UserControl.labRootItem.UBound Then
        PopupIndex = 0
    End If
    Call labRootItem_MouseDown(PopupIndex, vbLeftButton, 0, 0, 0)
    Call frmPopupMenu.Form_KeyDown(vbKeyDown, 0)
End Function

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Function GetLevels() As Integer()
    GetLevels = Levels
End Function

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub SetLevels(NewLevels() As Integer)
    Levels = NewLevels
    PropertyChanged
End Sub

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Function SetMenuCount(NewCount As Integer)
    ReDim Menus(NewCount)
    PropertyChanged
End Function

'Please note that this function is for internal usage only and is NOT suggested to call directly
Public Sub RaiseClickEvent(ByVal MenuID As Integer)
    RaiseEvent MenuItemClicked(MenuID)
End Sub

'Please note that this function is for internal usage only and is NOT suggested to call directly
'Please use menu properties instead
Public Sub SetMenuItemInfo(Index As Integer, MenuID As Integer, MenuText As String, Enabled As Boolean, CheckBox As Boolean, _
    Visible As Boolean, SubMenus() As String, SubMenuID() As Integer, MenuIcon() As Byte)
    
    With Menus(Index)
         .MenuID = MenuID
         .MenuText = MenuText
         .Enabled = Enabled
         .CheckBox = CheckBox
         .Visible = Visible
         .SubMenus = SubMenus
         .SubMenuID = SubMenuID
         .MenuIcon = MenuIcon
    End With
    PropertyChanged
    
    Call UpdateRootItems
End Sub

'Please note that this function is for internal usage only and is NOT suggested to call directly
'Please use menu properties instead
Public Sub GetMenuItemInfo(ByVal Index As Integer, MenuID As Integer, MenuText As String, Enabled As Boolean, CheckBox As Boolean, _
    Visible As Boolean, SubMenus() As String, SubMenuID() As Integer, MenuIcon() As Byte, Optional Checked As Boolean)
    
    With Menus(Index)
        MenuID = .MenuID
        MenuText = .MenuText
        Enabled = .Enabled
        CheckBox = .CheckBox
        Visible = .Visible
        SubMenus = .SubMenus
        SubMenuID = .SubMenuID
        MenuIcon = .MenuIcon
        If Not IsMissing(Checked) Then
            Checked = .Checked
        End If
    End With
End Sub

'Please note that the parameter 'HideFromSubMenu' is for internal usage ONLY. Please ignore it when calling this function
Public Sub HideMenu(Optional HideFromSubMenu As Boolean = False)
    Dim i           As Integer
    
    frmPopupMenu.CloseMenu
    PrevShowMenu = -1
    bShow = False
    If Not HideFromSubMenu Then
        For i = 0 To UserControl.labRootItem.UBound
            UserControl.labRootItem(i).BackColor = RGB(45, 45, 48)
        Next i
        UserControl.lnBorderLeft.Visible = False
        UserControl.lnBorderRight.Visible = False
        UserControl.lnBorderTop.Visible = False
    End If
End Sub

'Please note that the parameter 'ClickedFromRootItem' is for internal usage ONLY. Please ignore it when calling this function
Public Sub PopupMenu(ParentMenuID As Integer, Optional X As Single = -1, Optional Y As Single = -1, Optional ClickedFromRootItem As Integer = -1)
    On Error Resume Next
    
    If UBound(Menus(ParentMenuID + 1).SubMenuID) > 0 Then
        If ClickedFromRootItem = -1 Then
            Call HideMenu
        Else
            PopupIndex = ClickedFromRootItem
            frmPopupMenu.IsLastMenu = True
        End If
        frmPopupMenu.CloseMenu
        ReleaseCapture
        With frmPopupMenu
            If X = -1 And Y = -1 Then
                Dim CurPos  As POINT
                
                GetCursorPos CurPos
                .Left = CurPos.X * Screen.TwipsPerPixelX
                .Top = CurPos.Y * Screen.TwipsPerPixelY
            Else
                .Left = X
                .Top = Y
            End If
            If ClickedFromRootItem = -1 Then
                .AddItems Me, Menus(ParentMenuID + 1).SubMenuID, 0
            Else
                .AddItems Me, Menus(ParentMenuID + 1).SubMenuID, UserControl.labRootItem(ClickedFromRootItem).Width
            End If
            .NoWhitelist = False
            .Show
            .SetFocus
        End With
    End If
End Sub

Public Function GetMenuCount() As Integer
    GetMenuCount = UBound(Menus)
End Function

Private Sub UpdateRootItems()
    Dim i   As Integer
    
    For i = 1 To UserControl.labRootItem.UBound
        Unload UserControl.labRootItem(i)
    Next i
    For i = 1 To UBound(Menus)
        If Levels(i) = 0 Then
            ReDim Preserve RootMenuID(UBound(RootMenuID) + 1)
            RootMenuID(UBound(RootMenuID) - 1) = i
            
            If Menus(i).Visible Then
                If i = 1 Then
                    UserControl.labRootItem(0).Caption = String(SpaceCount, " ") & Menus(i).MenuText & String(SpaceCount, " ")
                    UserControl.labRootItem(0).Visible = True
                    UserControl.labRootItem(0).Left = 0
                    UserControl.labRootItem(0).AutoSize = True
                    UserControl.labRootItem(0).Enabled = Menus(i).Enabled
                Else
                    Load UserControl.labRootItem(UserControl.labRootItem.UBound + 1)
                    UserControl.labRootItem(UserControl.labRootItem.UBound).Caption = String(SpaceCount, " ") & Menus(i).MenuText & String(SpaceCount, " ")
                    UserControl.labRootItem(UserControl.labRootItem.UBound).Visible = True
                    UserControl.labRootItem(UserControl.labRootItem.UBound).Left = UserControl.labRootItem(UserControl.labRootItem.UBound - 1).Left + _
                        UserControl.labRootItem(UserControl.labRootItem.UBound - 1).Width
                    UserControl.labRootItem(UserControl.labRootItem.UBound).Enabled = Menus(i).Enabled
                End If
            End If
        End If
    Next i
    UserControl.Height = UserControl.labRootItem(0).Height + Me.HeightAddition
    
    For i = 0 To UserControl.labRootItem.UBound
        UserControl.labRootItem(i).Height = UserControl.Height
    Next i
End Sub

Private Sub labRootItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i       As Integer
    
    If Button <> vbLeftButton Then
        Exit Sub
    End If
    If bShow Then
        Call HideMenu
        Exit Sub
    Else
        bShow = True
        labRootItem(Index).BackColor = RGB(27, 27, 28)
        UserControl.lnBorderLeft.Visible = True
        UserControl.lnBorderRight.Visible = True
        UserControl.lnBorderTop.Visible = True
    End If
    '----------------------------------------------
    Dim wRect   As RECT
    
    If UBound(Menus(RootMenuID(Index)).SubMenus) > 0 Then
        GetWindowRect UserControl.hwnd, wRect
        With UserControl.lnBorderLeft
            .X1 = UserControl.labRootItem(Index).Left
            .Y1 = 0
            .X2 = UserControl.labRootItem(Index).Left
            .Y2 = UserControl.Height
        End With
        With UserControl.lnBorderTop
            .X1 = UserControl.labRootItem(Index).Left
            .Y1 = 0
            .X2 = UserControl.labRootItem(Index).Left + UserControl.labRootItem(Index).Width
            .Y2 = 0
        End With
        With UserControl.lnBorderRight
            .X1 = UserControl.labRootItem(Index).Left + UserControl.labRootItem(Index).Width
            .Y1 = 0
            .X2 = UserControl.labRootItem(Index).Left + UserControl.labRootItem(Index).Width
            .Y2 = UserControl.Height
        End With
        Call PopupMenu(RootMenuID(Index) - 1, wRect.Left * Screen.TwipsPerPixelX + UserControl.labRootItem(Index).Left, _
            wRect.bottom * Screen.TwipsPerPixelY, Index)
        PrevShowMenu = Index
    Else
        Call HideMenu(True)
        PrevShowMenu = Index
    End If
End Sub

Private Sub labRootItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i           As Label
    
    If Abs(PrevX - X) > 1 Or Abs(PrevY - Y) > 1 Then
        PrevX = X
        PrevY = Y
        For Each i In UserControl.labRootItem
            If i.Index <> Index Then
                i.BackColor = RGB(45, 45, 48)
            End If
        Next i
        
        If bShow Then
            If PrevShowMenu <> Index Then
                bShow = False
                Call labRootItem_MouseDown(Index, 1, 0, 0, 0)
            End If
        Else
            UserControl.labRootItem(Index).BackColor = RGB(92, 92, 94)
        End If
        UserControl.tmrCheckFocus.Enabled = True
    End If
End Sub

Private Sub labRootItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And UBound(Menus(RootMenuID(Index)).SubMenus) = 0 Then
        RaiseEvent MenuItemClicked(Menus(RootMenuID(Index)).MenuID)
        UserControl.labRootItem(Index).BackColor = RGB(45, 45, 48)
        bShow = False
    End If
End Sub

Private Sub tmrCheckFocus_Timer()
    Dim pt          As POINT
    
    GetCursorPos pt
    If Not bShow And WindowFromPoint(pt.X, pt.Y) <> UserControl.hwnd Then
        Call HideMenu
        Call UserControl_MouseMove(0, 0, 0, 0)
        UserControl.tmrCheckFocus.Enabled = False
    ElseIf bShow Then
        Dim wPopupMenu  As Form
        
        For Each wPopupMenu In Forms
            If wPopupMenu.Name = "frmPopupMenu" Then
                Exit Sub
            End If
        Next wPopupMenu
        Call HideMenu(True)
    End If
End Sub

Private Sub UserControl_Initialize()
    ReDim Menus(0)
    ReDim Menus(0).SubMenuID(0)
    ReDim Menus(0).SubMenus(0)
    ReDim Levels(0)
    ReDim RootMenuID(0)
    PrevShowMenu = -1
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then
        Call MoveLeft
    ElseIf KeyCode = vbKeyRight Then
        Call MoveRight
    ElseIf KeyCode = vbKeyReturn Then
        Call labRootItem_MouseUp(PopupIndex, vbLeftButton, 0, 0, 0)
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.HideMenu True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i           As Label
    
    If Not bShow Then
        For Each i In UserControl.labRootItem
            i.BackColor = RGB(45, 45, 48)
        Next i
    End If
    UserControl.tmrCheckFocus.Enabled = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Dim i           As Integer
    Dim j           As Integer
    
    Set labRootItem(0).Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_SpaceCount = PropBag.ReadProperty("SpaceCount", m_def_SpaceCount)
    m_HeightAddition = PropBag.ReadProperty("HeightAddition", m_def_HeightAddition)
    m_Transparent = PropBag.ReadProperty("Transparent", m_def_Transparent)
    
    ReDim Menus(PropBag.ReadProperty("MENU_ITEM_COUNT", 0))
    For i = 1 To UBound(Menus)
        With Menus(i)
            .MenuID = PropBag.ReadProperty("MenuID_" & i, i)
            .MenuText = PropBag.ReadProperty("MenuText_" & i, "")
            .Enabled = PropBag.ReadProperty("MenuEnabled_" & i, True)
            .CheckBox = PropBag.ReadProperty("MenuCheckBox_" & i, False)
            .Visible = PropBag.ReadProperty("MenuVisible_" & i, False)
            .MenuIcon = PropBag.ReadProperty("MenuIcon_" & i, StrConv("", vbFromUnicode))
            ReDim .SubMenuID(PropBag.ReadProperty("SUBMENU_ITEM_COUNT_" & i, 0))
            ReDim .SubMenus(PropBag.ReadProperty("SUBMENU_ITEM_COUNT_" & i, 0))
            For j = 0 To UBound(Menus(i).SubMenus)
                .SubMenus(j) = PropBag.ReadProperty("SubMenuText_" & i & "_" & j, "")
                .SubMenuID(j) = PropBag.ReadProperty("SubMenuID_" & i & "_" & j, j)
            Next j
        End With
    Next i
    ReDim Levels(PropBag.ReadProperty("LEVELS_COUNT", 0))
    For i = 0 To UBound(Levels)
        Levels(i) = PropBag.ReadProperty("LEVELS_" & i, 0)
    Next i
    
    Call UpdateRootItems
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Dim i           As Label
    
    UserControl.labRootItem(0).AutoSize = True
    UserControl.Height = UserControl.labRootItem(0).Height + Me.HeightAddition
    For Each i In UserControl.labRootItem
        i.Height = UserControl.Height
    Next i
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    Dim i           As Integer
    Dim j           As Integer
    
    Call PropBag.WriteProperty("Font", labRootItem(0).Font, Ambient.Font)
    Call PropBag.WriteProperty("SpaceCount", m_SpaceCount, m_def_SpaceCount)
    Call PropBag.WriteProperty("HeightAddition", m_HeightAddition, m_def_HeightAddition)
    Call PropBag.WriteProperty("Transparent", m_Transparent, m_def_Transparent)
    
    PropBag.WriteProperty "MENU_ITEM_COUNT", UBound(Menus), 0
    PropBag.WriteProperty "LEVELS_COUNT", UBound(Levels), 0
    For i = 0 To UBound(Levels)
        PropBag.WriteProperty "LEVELS_" & i, Levels(i), 0
    Next i
    With PropBag
        For i = 1 To UBound(Menus)
            .WriteProperty "MenuID_" & i, Menus(i).MenuID
            .WriteProperty "MenuText_" & i, Menus(i).MenuText, ""
            .WriteProperty "MenuEnabled_" & i, Menus(i).Enabled, True
            .WriteProperty "MenuCheckBox_" & i, Menus(i).CheckBox, False
            .WriteProperty "MenuVisible_" & i, Menus(i).Visible, False
            .WriteProperty "MenuIcon_" & i, Menus(i).MenuIcon, Nothing
            .WriteProperty "SUBMENU_ITEM_COUNT_" & i, UBound(Menus(i).SubMenus), 0
            For j = 0 To UBound(Menus(i).SubMenus)
                .WriteProperty "SubMenuText_" & i & "_" & j, Menus(i).SubMenus(j), ""
                .WriteProperty "SubMenuID_" & i & "_" & j, Menus(i).SubMenuID(j)
            Next j
        Next i
    End With
End Sub

Public Property Let MenuChecked(ByVal MenuID As Integer, ByVal bChecked As Boolean)
    Menus(MenuID + 1).Checked = bChecked
    Call UpdateRootItems
End Property

Public Property Get MenuChecked(ByVal MenuID As Integer) As Boolean
    MenuChecked = Menus(MenuID + 1).Checked
End Property

Public Property Let MenuText(ByVal MenuID As Integer, ByVal MenuText As String)
    Menus(MenuID + 1).MenuText = MenuText
    Call UpdateRootItems
End Property

Public Property Get MenuText(ByVal MenuID As Integer) As String
    MenuText = Menus(MenuID + 1).MenuText
End Property

Public Property Let MenuEnabled(ByVal MenuID As Integer, ByVal bEnabled As Boolean)
    Menus(MenuID + 1).Enabled = bEnabled
    Call UpdateRootItems
End Property

Public Property Get MenuEnabled(ByVal MenuID As Integer) As Boolean
    MenuEnabled = Menus(MenuID + 1).Enabled
End Property

Public Property Let MenuHasCheckBox(ByVal MenuID As Integer, ByVal bHasCheckbox As Boolean)
    Menus(MenuID + 1).CheckBox = bHasCheckbox
    Call UpdateRootItems
End Property

Public Property Get MenuHasCheckBox(ByVal MenuID As Integer) As Boolean
    MenuHasCheckBox = Menus(MenuID + 1).CheckBox
End Property

Public Property Let MenuVisible(ByVal MenuID As Integer, ByVal bVisible As Boolean)
    Menus(MenuID + 1).Visible = bVisible
    Call UpdateRootItems
End Property

Public Property Get MenuVisible(ByVal MenuID As Integer) As Boolean
    MenuVisible = Menus(MenuID + 1).Visible
End Property

Public Property Let MenuIcon(ByVal MenuID As Integer, NewIcon() As Byte)
    Menus(MenuID + 1).MenuIcon = NewIcon
End Property

Public Property Get MenuIcon(ByVal MenuID As Integer) As Byte()
    MenuIcon = Menus(MenuID).MenuIcon
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=labRootItem(0),labRootItem,0,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = labRootItem(0).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set labRootItem(0).Font = New_Font
    PropertyChanged "Font"
    
    Call UpdateRootItems
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,2
Public Property Get SpaceCount() As Integer
Attribute SpaceCount.VB_Description = "Returns/Sets the amount of spaces that will be added before and after the menu text. Use this property to adjust the position of the text."
    SpaceCount = m_SpaceCount
End Property

Public Property Let SpaceCount(ByVal New_SpaceCount As Integer)
    m_SpaceCount = New_SpaceCount
    PropertyChanged "SpaceCount"
    
    Call UpdateRootItems
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get HeightAddition() As Integer
Attribute HeightAddition.VB_Description = "Returns/Sets the addition to the height. Use this property if the control can not display the items properly."
    HeightAddition = m_HeightAddition
End Property

Public Property Let HeightAddition(ByVal New_HeightAddition As Integer)
    m_HeightAddition = New_HeightAddition
    PropertyChanged "HeightAddition"
    
    Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Transparent() As Boolean
Attribute Transparent.VB_Description = "Returns/Sets if the parent menu will be made transparent automatically if the sub-menu is displayed."
    Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal New_Transparent As Boolean)
    m_Transparent = New_Transparent
    PropertyChanged "Transparent"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SpaceCount = m_def_SpaceCount
    m_HeightAddition = m_def_HeightAddition
    m_Transparent = m_def_Transparent
End Sub

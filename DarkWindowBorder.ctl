VERSION 5.00
Begin VB.UserControl DarkWindowBorder 
   BackColor       =   &H00CC7A00&
   CanGetFocus     =   0   'False
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   540
   ScaleWidth      =   510
   ToolboxBitmap   =   "DarkWindowBorder.ctx":0000
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   0
      Picture         =   "DarkWindowBorder.ctx":0312
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "DarkWindowBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark°·WindowBorder by IceLolly
'Date: 2018.8.8

'               R    G   B
'Focused:       104, 33, 142
'No focus:      67,  67, 70
'Debugging:     202, 81, 0

'Please note that you should NOT use more than one of this
'control for the same window or it may cause unexpected bugs.

Dim BorderWindows(3)    As frmBorderWindow

Const m_def_Bind = 1
'Default Property Values:
Const m_def_UseSetParent = True
Const m_def_Transparency = 255
Const m_def_MinWidth = 0
Const m_def_MinHeight = 0
'Const m_def_MaxWidth = 0
'Const m_def_MaxHeight = 0
Const m_def_FocusedColor = &H8E2168             'RGB(104, 33, 142)
Const m_def_NotFocusedColor = &H454242          'RGB(66, 66, 69)
Const m_def_Thickness = 1
Const m_def_Sizable = 1
'Property Variables:
Dim m_UseSetParent As Boolean
Dim m_Transparency As Byte
Dim m_MinWidth As Long
Dim m_MinHeight As Long
Dim m_FocusedColor As OLE_COLOR
Dim m_NotFocusedColor As OLE_COLOR
Dim m_Thickness As Integer
Dim m_Sizable As Boolean
Dim m_Bind As Boolean

Private Sub RefreshState()
    On Error Resume Next
    Dim i   As Integer
    
    If Not Ambient.UserMode Then
        Exit Sub
    End If
    
    If m_Bind Then
        For i = 0 To 3
            Unload BorderWindows(i)
            
            Set BorderWindows(i) = New frmBorderWindow
            BorderWindows(i).Thickness = Thickness
            BorderWindows(i).BindPos = i
            BorderWindows(i).BoundWindow = UserControl.Parent.hWnd
            BorderWindows(i).fColor = Me.FocusedColor
            BorderWindows(i).nfColor = Me.NotFocusedColor
            BorderWindows(i).CanSize = Me.Sizable
            BorderWindows(i).MinH = Me.MinHeight
            BorderWindows(i).MinW = Me.MinWidth
            If Me.UseSetParent Then
                SetParent BorderWindows(i).hWnd, UserControl.Parent.hWnd
            End If
            BorderWindows(i).UseSetParent = Me.UseSetParent
            BorderWindows(i).Transparency = Me.Transparency
            BorderWindows(i).Show
        Next i
    Else
        For i = 0 To 3
            Unload BorderWindows(i)
        Next i
    End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Bind = m_def_Bind
    m_Sizable = m_def_Sizable
    m_Thickness = m_def_Thickness
    m_FocusedColor = m_def_FocusedColor
    m_NotFocusedColor = m_def_NotFocusedColor
    m_MinWidth = m_def_MinWidth
    m_MinHeight = m_def_MinHeight
    m_Transparency = m_def_Transparency
    m_UseSetParent = m_def_UseSetParent
    
    Call RefreshState
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Bind = PropBag.ReadProperty("Bind", m_def_Bind)
    m_Sizable = PropBag.ReadProperty("Sizable", m_def_Sizable)
    m_Thickness = PropBag.ReadProperty("Thickness", m_def_Thickness)
    m_FocusedColor = PropBag.ReadProperty("FocusedColor", m_def_FocusedColor)
    m_NotFocusedColor = PropBag.ReadProperty("NotFocusedColor", m_def_NotFocusedColor)
    m_MinWidth = PropBag.ReadProperty("MinWidth", m_def_MinWidth)
    m_MinHeight = PropBag.ReadProperty("MinHeight", m_def_MinHeight)
    m_Transparency = PropBag.ReadProperty("Transparency", m_def_Transparency)
    m_UseSetParent = PropBag.ReadProperty("UseSetParent", m_def_UseSetParent)
    
    Call RefreshState
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = UserControl.imgIcon.Width
    UserControl.Height = UserControl.imgIcon.Height
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    Dim i   As Integer
    
    For i = 0 To 3
        Unload BorderWindows(i)
    Next i
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Bind", m_Bind, m_def_Bind)
    Call PropBag.WriteProperty("Sizable", m_Sizable, m_def_Sizable)
    Call PropBag.WriteProperty("Thickness", m_Thickness, m_def_Thickness)
    Call PropBag.WriteProperty("FocusedColor", m_FocusedColor, m_def_FocusedColor)
    Call PropBag.WriteProperty("NotFocusedColor", m_NotFocusedColor, m_def_NotFocusedColor)
    Call PropBag.WriteProperty("MinWidth", m_MinWidth, m_def_MinWidth)
    Call PropBag.WriteProperty("MinHeight", m_MinHeight, m_def_MinHeight)
    Call PropBag.WriteProperty("Transparency", m_Transparency, m_def_Transparency)
    Call PropBag.WriteProperty("UseSetParent", m_UseSetParent, m_def_UseSetParent)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get Bind() As Boolean
Attribute Bind.VB_Description = "Return/sets if the border is bound with the parent window of this control."
    Bind = m_Bind
End Property

Public Property Let Bind(ByVal New_Bind As Boolean)
    m_Bind = New_Bind
    PropertyChanged "Bind"
    
    Call RefreshState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get Sizable() As Boolean
Attribute Sizable.VB_Description = "Returns/sets if the user can change the size of the bound window via border."
    Sizable = m_Sizable
End Property

Public Property Let Sizable(ByVal New_Sizable As Boolean)
    m_Sizable = New_Sizable
    PropertyChanged "Sizable"
    
    Call RefreshState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get Thickness() As Integer
Attribute Thickness.VB_Description = "Returns/sets the thickness of the border."
    Thickness = m_Thickness
End Property

Public Property Let Thickness(ByVal New_Thickness As Integer)
    m_Thickness = New_Thickness
    PropertyChanged "Thickness"
    
    Call RefreshState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8E2168
Public Property Get FocusedColor() As OLE_COLOR
Attribute FocusedColor.VB_Description = "Returns/sets the color of the border when the window is focused."
    FocusedColor = m_FocusedColor
End Property

Public Property Let FocusedColor(ByVal New_FocusedColor As OLE_COLOR)
    m_FocusedColor = New_FocusedColor
    PropertyChanged "FocusedColor"
    
    Call RefreshState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H464343
Public Property Get NotFocusedColor() As OLE_COLOR
Attribute NotFocusedColor.VB_Description = "Returns/sets the color of the border when the window isn't focused."
    NotFocusedColor = m_NotFocusedColor
End Property

Public Property Let NotFocusedColor(ByVal New_NotFocusedColor As OLE_COLOR)
    m_NotFocusedColor = New_NotFocusedColor
    PropertyChanged "NotFocusedColor"
    
    Call RefreshState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MinWidth() As Long
Attribute MinWidth.VB_Description = "Returns/Sets the minimum width (in pixels) can be changed via the border. 0 means not limited."
    MinWidth = m_MinWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Long)
    m_MinWidth = New_MinWidth
    PropertyChanged "MinWidth"
    
    Call RefreshState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MinHeight() As Long
Attribute MinHeight.VB_Description = "Returns/Sets the minimum height (in pixels) can be changed via the border. 0 means not limited."
    MinHeight = m_MinHeight
End Property

Public Property Let MinHeight(ByVal New_MinHeight As Long)
    m_MinHeight = New_MinHeight
    PropertyChanged "MinHeight"
    
    Call RefreshState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,255
Public Property Get Transparency() As Byte
Attribute Transparency.VB_Description = "Returns/Sets the transparency of the border window. 0 means invisible and the user can't interact with the border."
    Transparency = m_Transparency
End Property

Public Property Let Transparency(ByVal New_Transparency As Byte)
    m_Transparency = New_Transparency
    PropertyChanged "Transparency"
    
    Call RefreshState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get UseSetParent() As Boolean
Attribute UseSetParent.VB_Description = "Returns/Sets if the program use SetParent() or not. Please note that the border can't be transparent if uses SetParent()."
    UseSetParent = m_UseSetParent
End Property

Public Property Let UseSetParent(ByVal New_UseSetParent As Boolean)
    m_UseSetParent = New_UseSetParent
    PropertyChanged "UseSetParent"
    
    Call RefreshState
End Property


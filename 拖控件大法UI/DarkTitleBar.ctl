VERSION 5.00
Begin VB.UserControl DarkTitleBar 
   Alignable       =   -1  'True
   BackColor       =   &H00302D2D&
   ClientHeight    =   528
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5376
   ScaleHeight     =   528
   ScaleWidth      =   5376
   ToolboxBitmap   =   "DarkTitleBar.ctx":0000
   Begin 拖控件大法UI.DarkMenu mnuPopup 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4890
      _ExtentY        =   550
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MENU_ITEM_COUNT =   6
      LEVELS_COUNT    =   6
      LEVELS_2        =   1
      LEVELS_3        =   1
      LEVELS_4        =   1
      LEVELS_5        =   1
      LEVELS_6        =   1
      MenuID_1        =   0
      MenuText_1      =   "Popup"
      MenuVisible_1   =   -1  'True
      SUBMENU_ITEM_COUNT_1=   5
      SubMenuID_1_0   =   0
      SubMenuText_1_1 =   "还原"
      SubMenuID_1_1   =   2
      SubMenuText_1_2 =   "最大化"
      SubMenuID_1_2   =   3
      SubMenuText_1_3 =   "最小化"
      SubMenuID_1_3   =   4
      SubMenuText_1_4 =   "-"
      SubMenuID_1_4   =   5
      SubMenuText_1_5 =   "关闭"
      SubMenuID_1_5   =   6
      MenuID_2        =   1
      MenuText_2      =   "还原"
      MenuEnabled_2   =   0   'False
      MenuVisible_2   =   -1  'True
      MenuIcon_2      =   "DarkTitleBar.ctx":0312
      SubMenuID_2_0   =   0
      MenuID_3        =   2
      MenuText_3      =   "最大化"
      MenuVisible_3   =   -1  'True
      MenuIcon_3      =   "DarkTitleBar.ctx":04D8
      SubMenuID_3_0   =   0
      MenuID_4        =   3
      MenuText_4      =   "最小化"
      MenuVisible_4   =   -1  'True
      MenuIcon_4      =   "DarkTitleBar.ctx":069E
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "-"
      MenuVisible_5   =   -1  'True
      SubMenuID_5_0   =   0
      MenuID_6        =   5
      MenuText_6      =   "关闭"
      MenuVisible_6   =   -1  'True
      MenuIcon_6      =   "DarkTitleBar.ctx":0864
      SubMenuID_6_0   =   0
   End
   Begin VB.Timer tmrCheckFocus 
      Interval        =   10
      Left            =   3240
      Top             =   240
   End
   Begin 拖控件大法UI.DarkImageButton cmdMin 
      Height          =   480
      Left            =   3840
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "最小化"
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Picture         =   "DarkTitleBar.ctx":0A2A
      Focusable       =   0   'False
      HasBorder       =   0   'False
   End
   Begin 拖控件大法UI.DarkImageButton cmdMax 
      Height          =   480
      Left            =   4320
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "最大化"
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Picture         =   "DarkTitleBar.ctx":1704
      Focusable       =   0   'False
      HasBorder       =   0   'False
   End
   Begin 拖控件大法UI.DarkImageButton cmdClose 
      Height          =   480
      Left            =   4800
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "关闭"
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Picture         =   "DarkTitleBar.ctx":23DE
      Focusable       =   0   'False
      HasBorder       =   0   'False
   End
   Begin VB.Image imgMax 
      Height          =   384
      Left            =   4320
      Picture         =   "DarkTitleBar.ctx":30B8
      Top             =   480
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgRestore 
      Height          =   384
      Left            =   4800
      Picture         =   "DarkTitleBar.ctx":3D82
      Top             =   480
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dark♂TitleBar"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DADADA&
      Height          =   225
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgIcon 
      Height          =   384
      Left            =   0
      Picture         =   "DarkTitleBar.ctx":4A4C
      Top             =   0
      Width           =   384
   End
End
Attribute VB_Name = "DarkTitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark♂TitleBar by IceLolly
'Date: 2018.8.7

'               R    G    B
'Focused:       241, 241, 213
'No focus:      117, 153, 136

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Private Const SZ_MARGIN = 30

'Default Property Values:
Const m_def_BindCaption = 0
'Property Variables:
Dim m_BindCaption As Boolean

Private Sub cmdClose_Click()
    Unload UserControl.Parent
End Sub

Private Sub cmdMax_Click()
    On Error Resume Next
    Dim wp          As WINDOWPLACEMENT
    Dim PrevStyle   As Long
    
    GetWindowPlacement UserControl.Parent.hWnd, wp
    If wp.ShowCmd = SW_MAXIMIZE Then
        ShowWindow UserControl.Parent.hWnd, SW_RESTORE
        UserControl.cmdMax.ToolTipText = "最大化"
        Set UserControl.cmdMax.Picture = UserControl.imgMax.Picture
    Else
        ShowWindow UserControl.Parent.hWnd, SW_MAXIMIZE
        UserControl.cmdMax.ToolTipText = "还原"
        Set UserControl.cmdMax.Picture = UserControl.imgRestore.Picture
    End If
End Sub

Private Sub cmdMin_Click()
    On Error Resume Next
    ShowWindow UserControl.Parent.hWnd, SW_MINIMIZE
End Sub

Private Sub imgIcon_DblClick()
    UserControl.mnuPopup.HideMenu
    Call cmdClose_Click
End Sub

Private Sub imgIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Or Button = vbRightButton Then
        Dim wRect   As RECT
        Dim wp  As WINDOWPLACEMENT
    
        GetWindowPlacement UserControl.Parent.hWnd, wp
        If wp.ShowCmd = SW_MAXIMIZE Then
            UserControl.mnuPopup.MenuEnabled(1) = True
            UserControl.mnuPopup.MenuEnabled(2) = False
        Else
            UserControl.mnuPopup.MenuEnabled(1) = False
            UserControl.mnuPopup.MenuEnabled(2) = True
        End If
        
        GetWindowRect UserControl.hWnd, wRect
        UserControl.mnuPopup.PopupMenu 0, wRect.Left * Screen.TwipsPerPixelX + X + 120, _
            wRect.Top * Screen.TwipsPerPixelY + Y + 120
    End If
End Sub

Private Sub labTip_DblClick()
    If UserControl.cmdMax.Enabled Then
        Call cmdMax_Click
    End If
End Sub

Private Sub labTip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, 0, 0, 0)
End Sub

Private Sub mnuPopup_MenuItemClicked(MenuID As Integer)
    Select Case MenuID
        Case 1, 2
            If UserControl.cmdMax.Enabled Then
                Call cmdMax_Click
            End If
        
        Case 3
            If UserControl.cmdMin.Enabled Then
                Call cmdMin_Click
            End If
        
        Case 5
            Call cmdClose_Click
        
    End Select
End Sub

Private Sub tmrCheckFocus_Timer()
    On Error Resume Next
    
    If Not Ambient.UserMode Then
        UserControl.tmrCheckFocus.Enabled = False
    End If
    UserControl.Width = UserControl.Parent.ScaleWidth
    If GetForegroundWindow() = UserControl.Parent.hWnd Then
        UserControl.labTip.ForeColor = RGB(218, 218, 232)
    Else
        UserControl.labTip.ForeColor = RGB(188, 188, 188)
    End If
    If Me.BindCaption Then
        UserControl.labTip.Caption = UserControl.Parent.Caption
    End If
End Sub

Private Sub UserControl_DblClick()
    If UserControl.cmdMax.Enabled Then
        Call cmdMax_Click
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessageA UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
    ElseIf Button = vbRightButton Then
        Call imgIcon_MouseDown(1, 0, X, Y)
    End If
End Sub

Private Sub UserControl_Initialize()
    Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    UserControl.imgIcon.Left = SZ_MARGIN
    UserControl.labTip.Left = UserControl.imgIcon.Left + UserControl.imgIcon.Width + SZ_MARGIN
    UserControl.labTip.Top = UserControl.Height / 2 - UserControl.labTip.Height / 2
    UserControl.cmdClose.Left = UserControl.Width - UserControl.cmdClose.Width
    UserControl.cmdMax.Left = UserControl.cmdClose.Left - UserControl.cmdMax.Width
    UserControl.cmdMin.Left = UserControl.cmdMax.Left - UserControl.cmdMin.Width
    UserControl.cmdClose.Height = UserControl.Height
    UserControl.cmdMax.Height = UserControl.Height
    UserControl.cmdMin.Height = UserControl.Height
    UserControl.imgIcon.Top = UserControl.Height / 2 - imgIcon.Height / 2
    UserControl.imgIcon.Left = UserControl.imgIcon.Top
    UserControl.labTip.Left = UserControl.imgIcon.Left * 2 + UserControl.imgIcon.Width
    
    '-------------------------------------------------
    Dim wp  As WINDOWPLACEMENT
    
    GetWindowPlacement UserControl.Parent.hWnd, wp
    If wp.ShowCmd = SW_MAXIMIZE Then
        UserControl.cmdMax.ToolTipText = "还原"
        Set UserControl.cmdMax.Picture = UserControl.imgRestore.Picture
    Else
        UserControl.cmdMax.ToolTipText = "最大化"
        Set UserControl.cmdMax.Picture = UserControl.imgMax.Picture
    End If
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
'MappingInfo=labTip,labTip,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = labTip.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set labTip.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=labTip,labTip,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = labTip.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    labTip.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdMax,cmdMax,-1,Enabled
Public Property Get MaxButtonEnabled() As Boolean
    MaxButtonEnabled = cmdMax.Enabled
End Property

Public Property Let MaxButtonEnabled(ByVal New_MaxButtonEnabled As Boolean)
    cmdMax.Enabled() = New_MaxButtonEnabled
    PropertyChanged "MaxButtonEnabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdMin,cmdMin,-1,Enabled
Public Property Get MinButtonEnabled() As Boolean
    MinButtonEnabled = cmdMin.Enabled
End Property

Public Property Let MinButtonEnabled(ByVal New_MinButtonEnabled As Boolean)
    cmdMin.Enabled() = New_MinButtonEnabled
    PropertyChanged "MinButtonEnabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdClose,cmdClose,-1,Enabled
Public Property Get CloseButtonEnabled() As Boolean
    CloseButtonEnabled = cmdClose.Enabled
End Property

Public Property Let CloseButtonEnabled(ByVal New_CloseButtonEnabled As Boolean)
    cmdClose.Enabled() = New_CloseButtonEnabled
    PropertyChanged "CloseButtonEnabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set labTip.Font = PropBag.ReadProperty("Font", Ambient.Font)
    labTip.Caption = PropBag.ReadProperty("Caption", "Dark♂TitleBar")
    cmdMax.Enabled = PropBag.ReadProperty("MaxButtonEnabled", True)
    cmdMin.Enabled = PropBag.ReadProperty("MinButtonEnabled", True)
    cmdClose.Enabled = PropBag.ReadProperty("CloseButtonEnabled", True)
    m_BindCaption = PropBag.ReadProperty("BindCaption", m_def_BindCaption)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    
    Call UserControl_Resize
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", labTip.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", labTip.Caption, "Dark♂TitleBar")
    Call PropBag.WriteProperty("MaxButtonEnabled", cmdMax.Enabled, True)
    Call PropBag.WriteProperty("MinButtonEnabled", cmdMin.Enabled, True)
    Call PropBag.WriteProperty("CloseButtonEnabled", cmdClose.Enabled, True)
    Call PropBag.WriteProperty("BindCaption", m_BindCaption, m_def_BindCaption)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get BindCaption() As Boolean
Attribute BindCaption.VB_Description = "Return/Sets if the title changes with the parent window automatically"
    BindCaption = m_BindCaption
End Property

Public Property Let BindCaption(ByVal New_BindCaption As Boolean)
    m_BindCaption = New_BindCaption
    PropertyChanged "BindCaption"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BindCaption = m_def_BindCaption
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgIcon,imgIcon,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = imgIcon.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set imgIcon.Picture = New_Picture
    PropertyChanged "Picture"
End Property


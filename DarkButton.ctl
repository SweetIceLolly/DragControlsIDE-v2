VERSION 5.00
Begin VB.UserControl DarkButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00302D2D&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1320
   BeginProperty Font 
      Name            =   "Microsoft YaHei UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   495
   ScaleWidth      =   1320
   ToolboxBitmap   =   "DarkButton.ctx":0000
   Begin VB.Timer tmrSetColor 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   240
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dark¡áButton"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
   Begin VB.Shape shpFocus 
      BorderColor     =   &H00D5F1F1&
      BorderStyle     =   3  'Dot
      Height          =   615
      Left            =   120
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "DarkButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark¡áButton by IceLolly
'Date: 2018.8.6

'               R   G   B
'Normal:        45, 45, 48
'Mouse in:      92, 92, 94
'Mouse down:    0, 122, 204

Private Const SZ_BORDER = 30

Private Const NORMAL_R = 45, NORMAL_G = 45, NORMAL_B = 48
Private Const MOUSEIN_R = 92, MOUSEIN_G = 92, MOUSEIN_B = 94
Private Const MOUSEDOWN_R = 0, MOUSEDOWN_G = 122, MOUSEDOWN_B = 204

Dim BackR       As Integer
Dim BackG       As Integer
Dim BackB       As Integer

Dim bDown       As Boolean

'Default Property Values:
Const m_def_Alignment = 1
Const m_def_HasBorder = True
Const m_def_Enabled = True
'Property Variables:
Dim m_Alignment As Integer
Dim m_Enabled As Boolean
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

Private Sub tmrSetColor_Timer()
    Dim pt      As POINT
    Dim Target  As Long
    
    If Not Enabled Then
        Exit Sub
    End If
    If bDown Then
        If GetAsyncKeyState(VK_LBUTTON) = 0 Then
            Call UserControl_MouseUp(0, 0, 0, 0)
        Else
            Exit Sub
        End If
    End If
    GetCursorPos pt
    Target = WindowFromPoint(pt.X, pt.Y)
    If Target = UserControl.hWnd Then
        BackR = BackR + (MOUSEIN_R - NORMAL_R) / 30
        BackG = BackG + (MOUSEIN_G - NORMAL_G) / 30
        BackB = BackB + (MOUSEIN_B - NORMAL_B) / 30
        If BackR > MOUSEIN_R Or BackG > MOUSEIN_G Or BackB > MOUSEIN_B Then
            BackR = MOUSEIN_R
            BackG = MOUSEIN_G
            BackB = MOUSEIN_B
        End If
    Else
        BackR = BackR - (MOUSEIN_R - NORMAL_R) / 30
        BackG = BackG - (MOUSEIN_G - NORMAL_G) / 30
        BackB = BackB - (MOUSEIN_B - NORMAL_B) / 30
        If BackR < NORMAL_R Or BackG < NORMAL_G Or BackB < NORMAL_B Then
            BackR = NORMAL_R
            BackG = NORMAL_G
            BackB = NORMAL_B
            UserControl.tmrSetColor.Enabled = False
        End If
    End If
    UserControl.BackColor = RGB(BackR, BackG, BackB)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call UserControl_MouseDown(1, 0, 0, 0)
    ElseIf KeyCode = vbKeyReturn Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        RaiseEvent Click
        UserControl.BackColor = RGB(NORMAL_R, NORMAL_G, NORMAL_B)
    End If
End Sub

Private Sub UserControl_GotFocus()
    UserControl.shpFocus.Visible = True
End Sub

Private Sub UserControl_LostFocus()
    UserControl.shpFocus.Visible = False
    UserControl.BackColor = RGB(NORMAL_R, NORMAL_G, NORMAL_B)
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    
    BackR = NORMAL_R
    BackG = NORMAL_G
    BackB = NORMAL_B
    UserControl.BackColor = RGB(BackR, BackG, BackB)
    
    Call UserControl_Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        bDown = True
        UserControl.BackColor = RGB(MOUSEDOWN_R, MOUSEDOWN_G, MOUSEDOWN_B)
    End If
End Sub

Private Sub labTip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, 0, 0, 0)
End Sub

Private Sub labTip_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, 0, 0, 0)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.tmrSetColor.Enabled = True
End Sub

Private Sub labTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.labTip.ToolTipText = Extender.ToolTipText
    Call UserControl_MouseMove(0, 0, 0, 0)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Dim pt      As POINT
        Dim Target  As Long
        
        GetCursorPos pt
        Target = WindowFromPoint(pt.X, pt.Y)
        If Target = UserControl.hWnd Then
            RaiseEvent Click
        End If
    End If
    UserControl.BackColor = RGB(MOUSEIN_R, MOUSEIN_G, MOUSEIN_B)
    bDown = False
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next

    UserControl.shpFocus.Left = SZ_BORDER
    UserControl.shpFocus.Top = SZ_BORDER
    UserControl.shpFocus.Width = UserControl.Width - UserControl.shpFocus.Left - SZ_BORDER
    UserControl.shpFocus.Height = UserControl.Height - UserControl.shpFocus.Top - SZ_BORDER
    
    UserControl.labTip.Top = UserControl.Height / 2 - UserControl.labTip.Height / 2
    Select Case m_Alignment
        Case 0                                                              'Left
            UserControl.labTip.Left = SZ_BORDER + 60
        
        Case 1                                                              'Center
            UserControl.labTip.Left = UserControl.Width / 2 - UserControl.labTip.Width / 2
        
        Case 2                                                              'Right
            UserControl.labTip.Left = UserControl.Width - SZ_BORDER - 60 - UserControl.labTip.Width
        
    End Select
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"

    UserControl.labTip.Enabled = New_Enabled
    UserControl.Enabled = New_Enabled
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=labTip,labTip,-1,Font
Public Property Get Font() As Font
    Set Font = labTip.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set labTip.Font = New_Font
    PropertyChanged "Font"

    Call UserControl_Resize
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
    
    Call UserControl_Resize
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_Alignment = m_def_Alignment
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    Set labTip.Font = PropBag.ReadProperty("Font", Ambient.Font)
    labTip.Caption = PropBag.ReadProperty("Caption", "Caption")
    UserControl.BorderStyle = IIf(PropBag.ReadProperty("HasBorder", True), 1, 0)
    
    UserControl.labTip.Enabled = Me.Enabled
    UserControl.Enabled = Me.Enabled
    Call UserControl_Resize
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", labTip.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", labTip.Caption, "Caption")
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("HasBorder", IIf(UserControl.BorderStyle = 1, True, False), True)
    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get HasBorder() As Boolean
    HasBorder = IIf(UserControl.BorderStyle = 1, True, False)
End Property

Public Property Let HasBorder(ByVal New_HasBorder As Boolean)
    UserControl.BorderStyle() = IIf(New_HasBorder, 1, 0)
    PropertyChanged "HasBorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"
    
    UserControl_Resize
End Property


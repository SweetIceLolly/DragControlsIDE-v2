VERSION 5.00
Begin VB.UserControl DarkCheckBox 
   BackColor       =   &H00302D2D&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   ScaleHeight     =   375
   ScaleWidth      =   1980
   ToolboxBitmap   =   "DarkCheckBox.ctx":0000
   Begin VB.Timer tmrSetImage 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   360
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dark°·CheckBox"
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
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1410
   End
   Begin VB.Image imgYesMouseIn 
      Height          =   225
      Left            =   600
      Picture         =   "DarkCheckBox.ctx":0312
      Top             =   1200
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgYesMouseDown 
      Height          =   225
      Left            =   1080
      Picture         =   "DarkCheckBox.ctx":0668
      Top             =   1200
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgYesNormal 
      Height          =   225
      Left            =   120
      Picture         =   "DarkCheckBox.ctx":09BE
      Top             =   1200
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgNoMouseDown 
      Height          =   225
      Left            =   1080
      Picture         =   "DarkCheckBox.ctx":0D14
      Top             =   720
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgNoMouseIn 
      Height          =   225
      Left            =   600
      Picture         =   "DarkCheckBox.ctx":106A
      Top             =   720
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgNoNormal 
      Height          =   225
      Left            =   120
      Picture         =   "DarkCheckBox.ctx":13C0
      Top             =   720
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgState 
      Enabled         =   0   'False
      Height          =   225
      Left            =   120
      Picture         =   "DarkCheckBox.ctx":1716
      Top             =   120
      Width           =   225
   End
End
Attribute VB_Name = "DarkCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark°·CheckBox by IceLolly
'Date: 2018.8.6

Private Const SZ_DISTANCE = 60

Dim bDown       As Boolean
Dim bFocused    As Boolean

'Default Property Values:
Const m_def_Enabled = 0
Const m_def_Value = 0
'Property Variables:
Dim m_Value As Boolean
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

Private Sub labTip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, 0, 0, 0)
End Sub

Private Sub labTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.labTip.ToolTipText = Extender.ToolTipText
    Call UserControl_MouseMove(0, 0, 0, 0)
End Sub

Private Sub labTip_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, 0, 0, 0)
End Sub

Private Sub tmrSetImage_Timer()
    Dim pt      As POINT
    Dim Target  As Long
    
    GetCursorPos pt
    Target = WindowFromPoint(pt.X, pt.Y)
    If Target <> UserControl.hWnd Then
        If Me.Value Then
            UserControl.imgState.Picture = UserControl.imgYesNormal.Picture
        Else
            UserControl.imgState.Picture = UserControl.imgNoNormal.Picture
        End If
        UserControl.tmrSetImage.Enabled = False
    End If
End Sub

Private Sub UserControl_Initialize()
    Call UserControl_Resize
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        Call UserControl_MouseDown(1, 0, 0, 0)
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        Value = Not Value
        bDown = False
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_LostFocus()
    bFocused = False
    If Me.Value Then
        UserControl.imgState.Picture = UserControl.imgYesNormal.Picture
    Else
        UserControl.imgState.Picture = UserControl.imgNoNormal.Picture
    End If
End Sub

Private Sub UserControl_GotFocus()
    bFocused = True
    bDown = False
    If Me.Value = True Then
        UserControl.imgState.Picture = UserControl.imgYesMouseIn.Picture
    Else
        UserControl.imgState.Picture = UserControl.imgNoMouseIn.Picture
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        bDown = True
        If Me.Value Then
            UserControl.imgState.Picture = UserControl.imgYesMouseDown.Picture
        Else
            UserControl.imgState.Picture = UserControl.imgNoMouseDown.Picture
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDown = False
    If Button = vbLeftButton Then
        Dim pt      As POINT
        Dim Target  As Long
        
        GetCursorPos pt
        Target = WindowFromPoint(pt.X, pt.Y)
        If Target = UserControl.hWnd Then
            Value = Not Value
            RaiseEvent Click
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bDown Then
        If Me.Value = True Then
            UserControl.imgState.Picture = UserControl.imgYesMouseIn.Picture
        Else
            UserControl.imgState.Picture = UserControl.imgNoMouseIn.Picture
        End If
    Else
        If Me.Value = True Then
            UserControl.imgState.Picture = UserControl.imgYesMouseDown.Picture
        Else
            UserControl.imgState.Picture = UserControl.imgNoMouseDown.Picture
        End If
    End If
    UserControl.tmrSetImage.Enabled = True
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    UserControl.imgState.Left = SZ_DISTANCE
    UserControl.imgState.Top = UserControl.Height / 2 - UserControl.imgState.Height / 2
    UserControl.labTip.Top = UserControl.Height / 2 - UserControl.labTip.Height / 2
    UserControl.labTip.Left = UserControl.imgState.Left + UserControl.imgState.Width + SZ_DISTANCE
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set labTip.Font = PropBag.ReadProperty("Font", Ambient.Font)
    labTip.Caption = PropBag.ReadProperty("Caption", "Dark°·CheckBox")
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    
    Call UserControl_Resize
    UserControl.labTip.Enabled = UserControl.Enabled
    If Me.Value Then
        UserControl.imgState.Picture = UserControl.imgYesNormal.Picture
    Else
        UserControl.imgState.Picture = UserControl.imgNoNormal.Picture
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", labTip.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", labTip.Caption, "Dark°·CheckBox")
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
    
    UserControl.labTip.Enabled = New_Enabled
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
    Caption = labTip.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    labTip.Caption() = New_Caption
    PropertyChanged "Caption"
    
    Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    PropertyChanged "Value"
    
    If Not bFocused Then
        If New_Value Then
            UserControl.imgState.Picture = UserControl.imgYesNormal.Picture
        Else
            UserControl.imgState.Picture = UserControl.imgNoNormal.Picture
        End If
    Else
        If New_Value Then
            UserControl.imgState.Picture = UserControl.imgYesMouseIn.Picture
        Else
            UserControl.imgState.Picture = UserControl.imgNoMouseIn.Picture
        End If
    End If
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
End Sub

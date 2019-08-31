VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Image.ocx"
Begin VB.UserControl DarkImageButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00302D2D&
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2130
   ScaleHeight     =   570
   ScaleWidth      =   2130
   ToolboxBitmap   =   "DarkImageButton.ctx":0000
   Begin VB.Timer tmrSetColor 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   0
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dark°·Button"
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
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
   Begin ImageX.aicAlphaImage imgPicture 
      Height          =   480
      Left            =   120
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Image           =   "DarkImageButton.ctx":0312
      Enabled         =   0   'False
   End
   Begin VB.Shape shpFocus 
      BorderColor     =   &H00D5F1F1&
      BorderStyle     =   3  'Dot
      Height          =   495
      Left            =   120
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "DarkImageButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark°·ImageButton by IceLolly
'Date: 2018.8.7
'Actually, most of these codes are copied from Dark°·Button XD

'               R   G   B
'Normal:        45, 45, 48
'Mouse in:      92, 92, 94
'Mouse down:    0, 122, 204

Private Const SZ_BORDER = 30

Private Const NORMAL_R = 45, NORMAL_G = 45, NORMAL_B = 48
Private Const MOUSEIN_R = 92, MOUSEIN_G = 92, MOUSEIN_B = 94
Private Const MOUSEDOWN_R = 0, MOUSEDOWN_G = 122, MOUSEDOWN_B = 204

Private BackR   As Integer
Private BackG   As Integer
Private BackB   As Integer

Dim bDown       As Boolean

Dim imgData()   As Byte
Dim imgFileName As String

'Default Property Values:
Const m_def_Focusable = True
Const m_def_Enabled = True
'Property Variables:
Dim m_Focusable As Boolean
Dim m_Enabled As Boolean
Dim m_Alignment As Integer
'Event Declarations:
Event Click()

Private Sub labTip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, 0, 0, 0)
End Sub

Private Sub labTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, 0, 0, 0)
End Sub

Private Sub labTip_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, 0, 0, 0)
End Sub

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
    If Me.Focusable Then
        UserControl.shpFocus.Visible = True
    End If
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

Private Sub imgPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, 0, 0, 0)
End Sub

Private Sub imgPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, 0, 0, 0)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.tmrSetColor.Enabled = True
End Sub

Private Sub imgPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    UserControl.imgPicture.Height = UserControl.Height - SZ_BORDER * 2
    UserControl.imgPicture.Top = UserControl.Height / 2 - UserControl.imgPicture.Height / 2
    UserControl.imgPicture.Width = UserControl.imgPicture.Height
    
    Select Case Alignment
        Case 0
            UserControl.imgPicture.Left = SZ_BORDER * 2
            UserControl.labTip.Left = UserControl.imgPicture.Left + UserControl.imgPicture.Width + 120
        
        Case 1
            If UserControl.labTip.Caption <> "" Then
                UserControl.imgPicture.Left = UserControl.Width / 2 - (UserControl.imgPicture.Width + UserControl.labTip.Width + 120) / 2
            Else
                UserControl.imgPicture.Left = UserControl.Width / 2 - UserControl.imgPicture.Width / 2
            End If
            UserControl.labTip.Left = UserControl.imgPicture.Left + UserControl.imgPicture.Width + 120
        
        Case 2
            UserControl.labTip.Left = UserControl.Width - UserControl.labTip.Width - SZ_BORDER * 2
            UserControl.imgPicture.Left = labTip.Left - 120 - UserControl.imgPicture.Width
        
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

    UserControl.Enabled = New_Enabled
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_Focusable = m_def_Focusable
    m_Alignment = 1
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    imgData = PropBag.ReadProperty("Image", StrConv("", vbFromUnicode))
    m_Alignment = PropBag.ReadProperty("Alignment", 1)
    
    UserControl.imgPicture.LoadImage_FromArray imgData
    UserControl.Enabled = Me.Enabled
    UserControl.BorderStyle = IIf(PropBag.ReadProperty("HasBorder", True), 1, 0)
    m_Focusable = PropBag.ReadProperty("Focusable", m_def_Focusable)
    labTip.Caption = PropBag.ReadProperty("Caption", "Dark°·Button")
    imgPicture.AutoSize = PropBag.ReadProperty("AutoSize", False)
    
    Call UserControl_Resize
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Image", imgData, StrConv("", vbFromUnicode))
    Call PropBag.WriteProperty("Focusable", m_Focusable, m_def_Focusable)
    Call PropBag.WriteProperty("HasBorder", IIf(UserControl.BorderStyle = 1, True, False), True)
    Call PropBag.WriteProperty("Caption", labTip.Caption, "Dark°·Button")
    Call PropBag.WriteProperty("Alignment", m_Alignment, 1)
    Call PropBag.WriteProperty("AutoSize", imgPicture.AutoSize, False)
End Sub

Public Property Set Picture(ByVal New_Picture As Picture)
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    UserControl.imgPicture.LoadImage_FromStdPicture New_Picture
    PropertyChanged "Picture"
    
    Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Focusable() As Boolean
    Focusable = m_Focusable
End Property

Public Property Let Focusable(ByVal New_Focusable As Boolean)
    m_Focusable = New_Focusable
    PropertyChanged "Focusable"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get HasBorder() As Boolean
    HasBorder = IIf(UserControl.BorderStyle = 1, True, False)
End Property

Public Property Let HasBorder(ByVal New_HasBorder As Boolean)
    UserControl.BorderStyle() = IIf(New_HasBorder, 1, 0)
    PropertyChanged "HasBorder"
End Property

Public Property Get FileName() As String
    FileName = imgFileName
End Property

Public Property Let FileName(NewFileName As String)
    On Error Resume Next
    
    imgFileName = NewFileName
    UserControl.imgPicture.LoadImage_FromFile NewFileName
    Open NewFileName For Binary As #1
        ReDim imgData(LOF(1))
        Get #1, , imgData
    Close #1
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

Public Property Get Alignment() As Integer
    Alignment = m_Alignment
End Property

Public Property Let Alignment(New_Alignment As Integer)
    m_Alignment = New_Alignment
    Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgPicture,imgPicture,-1,AutoSize
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents"
    AutoSize = imgPicture.AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    imgPicture.AutoSize() = New_AutoSize
    PropertyChanged "AutoSize"
    Call UserControl_Resize
End Property


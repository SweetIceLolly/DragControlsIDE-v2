VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Image.ocx"
Begin VB.UserControl DarkImageButton 
   Appearance      =   0  'Flat
   BackColor       =   &H00302D2D&
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   ScaleHeight     =   465
   ScaleWidth      =   450
   ToolboxBitmap   =   "DarkImageButton.ctx":0000
   Begin VB.Timer tmrSetColor 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   360
   End
   Begin ImageX.aicAlphaImage imgPicture 
      Height          =   480
      Left            =   120
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Image           =   "DarkImageButton.ctx":0312
      Props           =   5
   End
   Begin VB.Shape shpFocus 
      BorderColor     =   &H00D5F1F1&
      BorderStyle     =   3  'Dot
      Height          =   495
      Left            =   120
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "DarkImageButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark¡áImageButton by IceLolly
'Date: 2018.8.7
'Actually, most of these codes are copied from Dark¡áButton XD

'               R   G   B
'Normal:        45, 45, 48
'Mouse in:      62, 62, 64
'Mouse down:    0, 122, 204

Private Const SZ_BORDER = 30

Private Const NORMAL_R = 45, NORMAL_G = 45, NORMAL_B = 48
Private Const MOUSEIN_R = 62, MOUSEIN_G = 62, MOUSEIN_B = 64
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
'Event Declarations:
Event Click()

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
        BackR = BackR - 1
        BackG = BackG - 1
        BackB = BackB - 1
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
    
    UserControl.imgPicture.Top = UserControl.Height / 2 - UserControl.imgPicture.Height / 2
    UserControl.imgPicture.Left = UserControl.Width / 2 - UserControl.imgPicture.Width / 2
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
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    imgData = PropBag.ReadProperty("Image", StrConv("", vbFromUnicode))
    
    UserControl.imgPicture.LoadImage_FromArray imgData
    UserControl.Enabled = Me.Enabled
    UserControl.BorderStyle = IIf(PropBag.ReadProperty("HasBorder", True), 1, 0)
    Call UserControl_Resize
    m_Focusable = PropBag.ReadProperty("Focusable", m_def_Focusable)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Image", imgData, StrConv("", vbFromUnicode))
    Call PropBag.WriteProperty("Focusable", m_Focusable, m_def_Focusable)
    Call PropBag.WriteProperty("HasBorder", IIf(UserControl.BorderStyle = 1, True, False), True)
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

VERSION 5.00
Begin VB.UserControl DarkEdit 
   BackColor       =   &H00463F3F&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   MousePointer    =   3  'I-Beam
   ScaleHeight     =   375
   ScaleWidth      =   1935
   ToolboxBitmap   =   "DarkEdit.ctx":0000
   Begin VB.Timer tmrSetColor 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   0
   End
   Begin VB.TextBox edMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00373333&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00899977&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "Boy°·Next°·Door"
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "DarkEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark°·Edit by IceLolly
'Date: 2018.8.9

'Border         R    G    B
'Normal:        63,  63,  70
'Mouse in:      0,   122, 204

'Edit back      R    G    B
'Normal:        51,  51,  55
'Mouse in:      63,  63,  70

'Edit text      R    G    B
'Normal:        119, 153, 137
'Mouse in:      255, 255, 255

Private Const SZ_BORDER = 10

Private Const BACK_NORMAL_R = 51, BACK_NORMAL_G = 51, BACK_NORMAL_B = 55
Private Const BACK_MOUSEIN_R = 63, BACK_MOUSEIN_G = 63, BACK_MOUSEIN_B = 70

Dim BackR       As Integer
Dim BackG       As Integer
Dim BackB       As Integer

Dim bFocused    As Boolean

'Event Declarations:
Event Change() 'MappingInfo=edMain,edMain,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Click() 'MappingInfo=edMain,edMain,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=edMain,edMain,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=edMain,edMain,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=edMain,edMain,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=edMain,edMain,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=edMain,edMain,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=edMain,edMain,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=edMain,edMain,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."

Private Sub edMain_GotFocus()
    bFocused = True
    UserControl.tmrSetColor.Enabled = True
    UserControl.BackColor = RGB(0, 122, 204)
    UserControl.edMain.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub edMain_LostFocus()
    bFocused = False
    UserControl.tmrSetColor.Enabled = False
    UserControl.BackColor = RGB(63, 63, 70)
    UserControl.edMain.ForeColor = RGB(119, 153, 137)
    UserControl.tmrSetColor.Enabled = True
End Sub

Private Sub tmrSetColor_Timer()
    On Error Resume Next
    Dim pt      As POINT
    Dim Target  As Long
    
    GetCursorPos pt
    Target = WindowFromPoint(pt.X, pt.Y)
    If GetForegroundWindow() <> UserControl.Parent.hWnd Then
        bFocused = False
    End If
    If GetFocus() = UserControl.edMain.hWnd Then
        bFocused = True
    End If
    If Target = UserControl.hWnd Or Target = UserControl.edMain.hWnd Or bFocused Then
        UserControl.BackColor = RGB(0, 122, 204)
        UserControl.edMain.ForeColor = RGB(255, 255, 255)
        BackR = BackR + 1
        BackG = BackG + 1
        BackB = BackB + 1
        If BackR > BACK_MOUSEIN_R Or BackG > BACK_MOUSEIN_G Or BackB > BACK_MOUSEIN_B Then
            BackR = BACK_MOUSEIN_R
            BackG = BACK_MOUSEIN_G
            BackB = BACK_MOUSEIN_B
        End If
    Else
        UserControl.BackColor = RGB(63, 63, 70)
        UserControl.edMain.ForeColor = RGB(119, 153, 137)
        BackR = BackR - 1
        BackG = BackG - 1
        BackB = BackB - 1
        If BackR < BACK_NORMAL_R Or BackG < BACK_NORMAL_G Or BackB < BACK_NORMAL_B Then
            BackR = BACK_NORMAL_R
            BackG = BACK_NORMAL_G
            BackB = BACK_NORMAL_B
        End If
    End If
    UserControl.edMain.BackColor = RGB(BackR, BackG, BackB)
End Sub

Private Sub edMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.edMain.ToolTipText = Extender.ToolTipText
    RaiseEvent MouseMove(Button, Shift, X, Y)
    Call UserControl_MouseMove(Button, 0, 0, 0)
End Sub

Private Sub UserControl_GotFocus()
    On Error Resume Next
    UserControl.edMain.SetFocus
    UserControl.tmrSetColor.Enabled = True
End Sub

Private Sub UserControl_Initialize()
    BackR = BACK_NORMAL_R
    BackG = BACK_NORMAL_G
    BackB = BACK_NORMAL_B
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.tmrSetColor.Enabled = True
    UserControl.BackColor = RGB(0, 122, 204)
    UserControl.edMain.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub UserControl_Resize()
    UserControl.edMain.Left = SZ_BORDER
    UserControl.edMain.Top = SZ_BORDER
    UserControl.edMain.Width = UserControl.Width - SZ_BORDER * 3
    UserControl.edMain.Height = UserControl.Height - SZ_BORDER * 3
End Sub
Private Sub edMain_Change()
    RaiseEvent Change
End Sub

Private Sub edMain_Click()
    RaiseEvent Click
End Sub

Private Sub edMain_DblClick()
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

Private Sub edMain_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub edMain_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub edMain_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=edMain,edMain,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = edMain.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    edMain.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=edMain,edMain,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = edMain.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    edMain.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=edMain,edMain,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
    MultiLine = edMain.MultiLine
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
'MappingInfo=edMain,edMain,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
    PasswordChar = edMain.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    edMain.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

Private Sub edMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub edMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set edMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    edMain.Locked = PropBag.ReadProperty("Locked", False)
    edMain.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    edMain.Text = PropBag.ReadProperty("Text", "Boy°·Next°·Door")
    edMain.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    edMain.SelLength = PropBag.ReadProperty("SelLength", 0)
    edMain.SelStart = PropBag.ReadProperty("SelStart", 0)
    edMain.SelText = PropBag.ReadProperty("SelText", "")
    edMain.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", edMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("Locked", edMain.Locked, False)
    Call PropBag.WriteProperty("MaxLength", edMain.MaxLength, 0)
    Call PropBag.WriteProperty("Text", edMain.Text, "Boy°·Next°·Door")
    Call PropBag.WriteProperty("PasswordChar", edMain.PasswordChar, "")
    Call PropBag.WriteProperty("SelLength", edMain.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", edMain.SelStart, 0)
    Call PropBag.WriteProperty("SelText", edMain.SelText, "")
    Call PropBag.WriteProperty("MousePointer", edMain.MousePointer, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=edMain,edMain,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = edMain.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    edMain.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=edMain,edMain,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = edMain.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    edMain.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=edMain,edMain,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = edMain.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    edMain.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=edMain,edMain,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = edMain.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    edMain.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property


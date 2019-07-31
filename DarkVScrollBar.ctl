VERSION 5.00
Begin VB.UserControl DarkVScrollBar 
   BackColor       =   &H00423E3E&
   CanGetFocus     =   0   'False
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   255
   ScaleHeight     =   2835
   ScaleWidth      =   255
   ToolboxBitmap   =   "DarkVScrollBar.ctx":0000
   Begin VB.Timer tmrCheckFocus 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   2280
   End
   Begin VB.Image imgUpMouseDown 
      Height          =   240
      Left            =   1200
      Picture         =   "DarkVScrollBar.ctx":0312
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgUpMouseIn 
      Height          =   240
      Left            =   840
      Picture         =   "DarkVScrollBar.ctx":069C
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgUpNormal 
      Height          =   240
      Left            =   480
      Picture         =   "DarkVScrollBar.ctx":0A26
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDownMouseDown 
      Height          =   240
      Left            =   1200
      Picture         =   "DarkVScrollBar.ctx":0DB0
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDownMouseIn 
      Height          =   240
      Left            =   840
      Picture         =   "DarkVScrollBar.ctx":113A
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDownNormal 
      Height          =   240
      Left            =   480
      Picture         =   "DarkVScrollBar.ctx":14C4
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape shpBar 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00686868&
      FillStyle       =   0  'Solid
      Height          =   1200
      Left            =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Image imgDown 
      Height          =   240
      Left            =   0
      Picture         =   "DarkVScrollBar.ctx":184E
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Left            =   0
      Picture         =   "DarkVScrollBar.ctx":1BD8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "DarkVScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark°·VScrollBar by IceLolly
'Date: 2018.8.10

'               R    G    B
'Back:          62,  62,  66

'Bar            R    G    B
'Normal:        104, 104, 104
'Mouse in:      158, 158, 158
'Mouse down:    239, 235, 239

Private Const BAR_MARGIN = 60

Dim DownPos     As Long
Dim DownY       As Single
Dim bDown       As Boolean
Dim bUpDown     As Boolean
Dim bDownDown   As Boolean
Dim baUpDown    As Boolean
Dim baDownDown  As Boolean
Dim TargetY     As Single
Dim DownTime    As Long

'Default Property Values:
Const m_def_BarHeight = 1200
Const m_def_Max = 100
Const m_def_Min = 0
Const m_def_SmallChange = 1
Const m_def_LargeChange = 5
Const m_def_Value = 0
'Property Variables:
Dim m_BarHeight As Long
Dim m_Max As Long
Dim m_Min As Long
Dim m_SmallChange As Long
Dim m_LargeChange As Long
Dim m_Value As Long
'Event Declarations:
Event ValueChanged(NewValue As Long)
Attribute ValueChanged.VB_Description = "Invoked when the value of the bar is changed."

Private Sub imgDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        bDownDown = True
        UserControl.imgDown.Picture = UserControl.imgDownMouseDown.Picture
        If Value < Max Then
            Value = Value + SmallChange
            If Value > Max Then
                Value = Max
            End If
        End If
        DownTime = GetTickCount
        UserControl.tmrCheckFocus.Enabled = True
    End If
End Sub

Private Sub imgDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bDownDown Then
        UserControl.imgDown.Picture = UserControl.imgDownMouseIn.Picture
        UserControl.tmrCheckFocus.Enabled = True
    End If
End Sub

Private Sub imgDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDownDown = False
End Sub

Private Sub imgUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        bUpDown = True
        UserControl.imgUp.Picture = UserControl.imgUpMouseDown.Picture
        If Value > Min Then
            Value = Value - SmallChange
            If Value < Min Then
                Value = Min
            End If
        End If
        DownTime = GetTickCount
        UserControl.tmrCheckFocus.Enabled = True
    End If
End Sub

Private Sub imgUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bUpDown Then
        UserControl.imgUp.Picture = UserControl.imgUpMouseIn.Picture
        UserControl.tmrCheckFocus.Enabled = True
    End If
End Sub

Private Sub imgUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bUpDown = False
End Sub

Private Sub tmrCheckFocus_Timer()
    On Error Resume Next
    
    Dim pt      As POINT
    Dim Target  As Long
    
    GetCursorPos pt
    Target = WindowFromPoint(pt.X, pt.Y)
    If GetAsyncKeyState(VK_LBUTTON) = 0 Then
        bDown = False
        bDownDown = False
        bUpDown = False
        baUpDown = False
        baDownDown = False
    End If
    If bDownDown And (GetTickCount - DownTime) > 500 Then
        If Value < Max Then
            Value = Value + SmallChange
            If Value > Max Then
                Value = Max
            End If
        End If
    ElseIf bUpDown And (GetTickCount - DownTime) > 500 Then
        If Value > Min Then
            Value = Value - SmallChange
            If Value < Min Then
                Value = Min
            End If
        End If
    ElseIf baDownDown And (GetTickCount - DownTime) > 500 Then
        If Value < Max And UserControl.shpBar.Top + UserControl.shpBar.Height < TargetY Then
            Value = Value + LargeChange
            If Value > Max Then
                Value = Max
            End If
        End If
    ElseIf baUpDown And (GetTickCount - DownTime) > 500 Then
        If Value > Min And UserControl.shpBar.Top > TargetY Then
            Value = Value - LargeChange
            If Value < Min Then
                Value = Min
            End If
        End If
    End If
    If bDown Then
        Dim NewPos  As Long
        Dim NewVal  As Long
        
        NewPos = DownY + (pt.Y - DownPos) * Screen.TwipsPerPixelY
        If NewPos < UserControl.imgUp.Height Then
            NewPos = UserControl.imgUp.Height
        ElseIf NewPos + UserControl.shpBar.Height > UserControl.imgDown.Top Then
            NewPos = UserControl.imgDown.Top - UserControl.shpBar.Height
        End If
        NewVal = Min + (Max - Min) / (UserControl.imgDown.Top - UserControl.shpBar.Height - UserControl.imgUp.Height) * (NewPos - UserControl.imgUp.Height)
        If Value <> NewVal Then
            Value = NewVal
        End If
    End If
    If Target <> UserControl.hWnd And Not bDown Then
        UserControl.shpBar.FillColor = RGB(104, 104, 104)
        UserControl.imgUp.Picture = UserControl.imgUpNormal.Picture
        UserControl.imgDown.Picture = UserControl.imgDownNormal.Picture
        UserControl.tmrCheckFocus.Enabled = False
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If Y >= UserControl.shpBar.Top And Y <= UserControl.shpBar.Top + UserControl.shpBar.Height Then
            Dim pt      As POINT
            
            GetCursorPos pt
            DownY = UserControl.shpBar.Top
            DownPos = pt.Y
            bDown = True
        ElseIf Y < UserControl.shpBar.Top Then
            TargetY = Y
            baUpDown = True
            If Value > Min Then
                Value = Value - LargeChange
                If Value < Min Then
                    Value = Min
                End If
            End If
            DownTime = GetTickCount
            UserControl.tmrCheckFocus.Enabled = True
        ElseIf Y > UserControl.shpBar.Top + UserControl.shpBar.Height Then
            TargetY = Y
            baDownDown = True
            If Value < Max Then
                Value = Value + LargeChange
                If Value > Max Then
                    Value = Max
                End If
            End If
            DownTime = GetTickCount
            UserControl.tmrCheckFocus.Enabled = True
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.imgUp.Picture = UserControl.imgUpNormal.Picture
    UserControl.imgDown.Picture = UserControl.imgDownNormal.Picture
    If baDownDown Or baUpDown Then
        TargetY = Y
    End If
    If Y >= UserControl.shpBar.Top And Y <= UserControl.shpBar.Top + UserControl.shpBar.Height Then
        UserControl.shpBar.FillColor = RGB(158, 158, 158)
        UserControl.tmrCheckFocus.Enabled = True
    ElseIf Not bDown Then
        UserControl.shpBar.FillColor = RGB(104, 104, 104)
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    UserControl.imgUp.Top = 0
    UserControl.imgUp.Left = 0
    UserControl.imgUp.Width = UserControl.Width
    UserControl.imgUp.Height = 240
    UserControl.imgDown.Width = UserControl.Width
    UserControl.imgDown.Height = 240
    UserControl.imgDown.Top = UserControl.Height - UserControl.imgDown.Height
    UserControl.imgDown.Left = 0
    UserControl.shpBar.Left = BAR_MARGIN
    UserControl.shpBar.Width = UserControl.Width - BAR_MARGIN * 2
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/Sets the maximum value of the bar."
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/Sets the minimum value of the bar."
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SmallChange() As Long
Attribute SmallChange.VB_Description = "Returns/Sets the change in value when the user clicks the change button."
    SmallChange = m_SmallChange
End Property

Public Property Let SmallChange(ByVal New_SmallChange As Long)
    m_SmallChange = New_SmallChange
    PropertyChanged "SmallChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get LargeChange() As Long
Attribute LargeChange.VB_Description = "Returns/Sets the change in value when the user clicks in the bar area."
    LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Long)
    m_LargeChange = New_LargeChange
    PropertyChanged "LargeChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/Sets the value of the bar."
    Value = m_Value
End Property

Public Property Let Value(New_Value As Long)
    If New_Value < Min Then
        New_Value = Min
    ElseIf New_Value > Max Then
        New_Value = Max
    End If
    m_Value = New_Value
    PropertyChanged "Value"
    UserControl.shpBar.Top = (New_Value - Min) * (UserControl.imgDown.Top - UserControl.shpBar.Height - _
        UserControl.imgUp.Height) / (Max - Min) + UserControl.imgUp.Height
    RaiseEvent ValueChanged(New_Value)
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

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_SmallChange = m_def_SmallChange
    m_LargeChange = m_def_LargeChange
    m_Value = m_def_Value
    m_BarHeight = m_def_BarHeight
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_SmallChange = PropBag.ReadProperty("SmallChange", m_def_SmallChange)
    m_LargeChange = PropBag.ReadProperty("LargeChange", m_def_LargeChange)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_BarHeight = PropBag.ReadProperty("BarHeight", m_def_BarHeight)
    
    UserControl.shpBar.Height = m_BarHeight
    If m_Value < Min Then
        m_Value = Min
    ElseIf m_Value > Max Then
        m_Value = Max
    End If
    Value = m_Value
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("SmallChange", m_SmallChange, m_def_SmallChange)
    Call PropBag.WriteProperty("LargeChange", m_LargeChange, m_def_LargeChange)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BarHeight", m_BarHeight, m_def_BarHeight)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1200
Public Property Get BarHeight() As Long
Attribute BarHeight.VB_Description = "Returns/sets the height of the bar."
    BarHeight = m_BarHeight
End Property

Public Property Let BarHeight(ByVal New_BarHeight As Long)
    On Error Resume Next
    m_BarHeight = New_BarHeight
    PropertyChanged "BarHeight"
    
    UserControl.shpBar.Height = New_BarHeight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property


VERSION 5.00
Begin VB.UserControl DarkHScrollBar 
   BackColor       =   &H00423E3E&
   CanGetFocus     =   0   'False
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   ScaleHeight     =   240
   ScaleWidth      =   3135
   ToolboxBitmap   =   "DarkHScrollBar.ctx":0000
   Begin VB.Timer tmrCheckFocus 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2040
      Top             =   360
   End
   Begin VB.Image imgLeft 
      Height          =   240
      Left            =   0
      Picture         =   "DarkHScrollBar.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgRight 
      Height          =   240
      Left            =   2760
      Picture         =   "DarkHScrollBar.ctx":069C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Shape shpBar 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00686868&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   240
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image imgLeftNormal 
      Height          =   240
      Left            =   480
      Picture         =   "DarkHScrollBar.ctx":0A26
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgLeftMouseIn 
      Height          =   240
      Left            =   840
      Picture         =   "DarkHScrollBar.ctx":0DB0
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgLeftMouseDown 
      Height          =   240
      Left            =   1200
      Picture         =   "DarkHScrollBar.ctx":113A
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRightNormal 
      Height          =   240
      Left            =   480
      Picture         =   "DarkHScrollBar.ctx":14C4
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRightMouseIn 
      Height          =   240
      Left            =   840
      Picture         =   "DarkHScrollBar.ctx":184E
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRightMouseDown 
      Height          =   240
      Left            =   1200
      Picture         =   "DarkHScrollBar.ctx":1BD8
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "DarkHScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dark°·HScrollBar by IceLolly
'Date: 2018.9.7

'               R    G    B
'Back:          62,  62,  66

'Bar            R    G    B
'Normal:        104, 104, 104
'Mouse in:      158, 158, 158
'Mouse down:    239, 235, 239

Private Const BAR_MARGIN = 60

Dim DownPos     As Long
Dim DownX       As Single
Dim bDown       As Boolean
Dim bLeftDown   As Boolean
Dim bRightDown  As Boolean
Dim baLeftDown  As Boolean
Dim baRightDown As Boolean
Dim TargetX     As Single
Dim DownTime    As Long

'Default Property Values:
Const m_def_BarWidth = 1200
Const m_def_Max = 100
Const m_def_Min = 0
Const m_def_SmallChange = 1
Const m_def_LargeChange = 5
Const m_def_Value = 0
'Property Variables:
Dim m_BarWidth As Long
Dim m_Max As Long
Dim m_Min As Long
Dim m_SmallChange As Long
Dim m_LargeChange As Long
Dim m_Value As Long
'Event Declarations:
Event ValueChanged(NewValue As Long)

Private Sub imgRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        bRightDown = True
        UserControl.imgRight.Picture = UserControl.imgRightMouseDown.Picture
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

Private Sub imgRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bRightDown Then
        UserControl.imgRight.Picture = UserControl.imgRightMouseIn.Picture
        UserControl.tmrCheckFocus.Enabled = True
    End If
End Sub

Private Sub imgRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bRightDown = False
End Sub

Private Sub imgLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        bLeftDown = True
        UserControl.imgLeft.Picture = UserControl.imgLeftMouseDown.Picture
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

Private Sub imgLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLeftDown Then
        UserControl.imgLeft.Picture = UserControl.imgLeftMouseIn.Picture
        UserControl.tmrCheckFocus.Enabled = True
    End If
End Sub

Private Sub imgLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bLeftDown = False
End Sub

Private Sub tmrCheckFocus_Timer()
    On Error Resume Next
    
    Dim pt          As POINT
    Dim Target      As Long
    
    GetCursorPos pt
    Target = WindowFromPoint(pt.X, pt.Y)
    If GetAsyncKeyState(VK_LBUTTON) = 0 Then
        bDown = False
        bLeftDown = False
        bRightDown = False
        baLeftDown = False
        baRightDown = False
    End If
    If bRightDown And (GetTickCount - DownTime) > 500 Then
        If Value < Max Then
            Value = Value + SmallChange
            If Value > Max Then
                Value = Max
            End If
        End If
    ElseIf bLeftDown And (GetTickCount - DownTime) > 500 Then
        If Value > Min Then
            Value = Value - SmallChange
            If Value < Min Then
                Value = Min
            End If
        End If
    ElseIf baRightDown And (GetTickCount - DownTime) > 500 Then
        If Value < Max And UserControl.shpBar.Left + UserControl.shpBar.Width < TargetX Then
            Value = Value + LargeChange
            If Value > Max Then
                Value = Max
            End If
        End If
    ElseIf baLeftDown And (GetTickCount - DownTime) > 500 Then
        If Value > Min And UserControl.shpBar.Left > TargetX Then
            Value = Value - LargeChange
            If Value < Min Then
                Value = Min
            End If
        End If
    End If
    If bDown Then
        Dim NewPos  As Long
        Dim NewVal  As Long
        
        NewPos = DownX + (pt.X - DownPos) * Screen.TwipsPerPixelX
        If NewPos < UserControl.imgLeft.Width Then
            NewPos = UserControl.imgLeft.Width
        ElseIf NewPos + UserControl.shpBar.Width > UserControl.imgRight.Left Then
            NewPos = UserControl.imgRight.Left - UserControl.shpBar.Width
        End If
        NewVal = Min + (Max - Min) / (UserControl.imgRight.Left - UserControl.shpBar.Width - UserControl.imgLeft.Width) * (NewPos - UserControl.imgLeft.Width)
        If Value <> NewVal Then
            Value = NewVal
        End If
    End If
    If Target <> UserControl.hWnd And Not bDown Then
        UserControl.shpBar.FillColor = RGB(104, 104, 104)
        UserControl.imgLeft.Picture = UserControl.imgLeftNormal.Picture
        UserControl.imgRight.Picture = UserControl.imgRightNormal.Picture
        UserControl.tmrCheckFocus.Enabled = False
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If X > UserControl.shpBar.Left And X < UserControl.shpBar.Left + UserControl.shpBar.Width Then
            Dim pt      As POINT
            
            GetCursorPos pt
            DownX = UserControl.shpBar.Left
            DownPos = pt.X
            bDown = True
        ElseIf X < UserControl.shpBar.Left Then
            TargetX = X
            baLeftDown = True
            If Value > Min Then
                Value = Value - LargeChange
                If Value < Min Then
                    Value = Min
                End If
            End If
            DownTime = GetTickCount
            UserControl.tmrCheckFocus.Enabled = True
        ElseIf X > UserControl.shpBar.Left + UserControl.shpBar.Width Then
            TargetX = X
            baRightDown = True
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
    UserControl.imgLeft.Picture = UserControl.imgLeftNormal.Picture
    UserControl.imgRight.Picture = UserControl.imgRightNormal.Picture
    If baLeftDown Or baRightDown Then
        TargetX = X
    End If
    If X >= UserControl.shpBar.Left And X <= UserControl.shpBar.Left + UserControl.shpBar.Width Then
        UserControl.shpBar.FillColor = RGB(158, 158, 158)
        UserControl.tmrCheckFocus.Enabled = True
    ElseIf Not bDown Then
        UserControl.shpBar.FillColor = RGB(104, 104, 104)
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    UserControl.imgLeft.Left = 0
    UserControl.imgLeft.Top = 0
    UserControl.imgLeft.Height = UserControl.Height
    UserControl.imgLeft.Width = 240
    UserControl.imgRight.Height = UserControl.Height
    UserControl.imgRight.Width = 240
    UserControl.imgRight.Left = UserControl.Width - UserControl.imgRight.Width
    UserControl.imgRight.Top = 0
    UserControl.shpBar.Top = BAR_MARGIN
    UserControl.shpBar.Height = UserControl.Height - BAR_MARGIN * 2
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Min() As Long
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SmallChange() As Long
    SmallChange = m_SmallChange
End Property

Public Property Let SmallChange(ByVal New_SmallChange As Long)
    m_SmallChange = New_SmallChange
    PropertyChanged "SmallChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get LargeChange() As Long
    LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Long)
    m_LargeChange = New_LargeChange
    PropertyChanged "LargeChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
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
    UserControl.shpBar.Left = (New_Value - Min) * (UserControl.imgRight.Left - UserControl.shpBar.Width - _
        UserControl.imgLeft.Width) / (Max - Min) + UserControl.imgLeft.Width
    RaiseEvent ValueChanged(New_Value)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
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
    m_BarWidth = m_def_BarWidth
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_SmallChange = PropBag.ReadProperty("SmallChange", m_def_SmallChange)
    m_LargeChange = PropBag.ReadProperty("LargeChange", m_def_LargeChange)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_BarWidth = PropBag.ReadProperty("BarWidth", m_def_BarWidth)
    
    UserControl.shpBar.Width = m_BarWidth
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
    Call PropBag.WriteProperty("BarWidth", m_BarWidth, m_def_BarWidth)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1200
Public Property Get BarWidth() As Long
    BarWidth = m_BarWidth
End Property

Public Property Let BarWidth(ByVal New_BarWidth As Long)
    m_BarWidth = New_BarWidth
    PropertyChanged "BarWidth"
    
    UserControl.shpBar.Width = New_BarWidth
End Property


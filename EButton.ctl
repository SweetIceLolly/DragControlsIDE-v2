VERSION 5.00
Begin VB.UserControl DarkEButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Animation 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4080
      Top             =   1200
   End
   Begin VB.Label hover 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2925
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Label con 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   1560
      Width           =   135
   End
End
Attribute VB_Name = "DarkEButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event Click()
Event hover(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Leave()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ndec As Long, nhoc As Long, cd As Boolean, inTime As Long, na As EAlign
Dim lc(3) As Byte, nc(3) As Byte, di As Integer
Public Enum EAlign
    onCenter = 0
    onLeft = 1
    onRight = 2
End Enum
Public Property Get align() As EAlign
    align = na
End Property
Public Property Let align(a As EAlign)
    na = a
    Call UserControl_Resize
End Property
Public Property Get DefaultColor() As OLE_COLOR
    DefaultColor = ndec
End Property
Public Property Let DefaultColor(c As OLE_COLOR)
    ndec = c
    UserControl.Backcolor = c
End Property
Public Property Get HoverColor() As OLE_COLOR
    HoverColor = nhoc
End Property
Public Property Let HoverColor(c As OLE_COLOR)
    nhoc = c
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = con.ForeColor
End Property
Public Property Let ForeColor(c As OLE_COLOR)
    con.ForeColor = c
End Property
Public Property Get Font() As StdFont
    Set Font = con.Font
End Property
Public Property Set Font(f As StdFont)
    Set con.Font = f
End Property
Public Property Get Content() As String
    Content = con.Caption
End Property
Public Property Let Content(c As String)
    con.Caption = c
    Call UserControl_Resize
End Property

Private Sub Animation_Timer()
    If GetTickCount - inTime <= 300 Then
        Dim pro As Single, buff(2) As Long
        pro = (GetTickCount - inTime) / 300
        buff(0) = nc(0): buff(0) = buff(0) - lc(0): buff(1) = nc(1): buff(1) = buff(1) - lc(1): buff(2) = nc(2): buff(2) = buff(2) - lc(2)
        UserControl.Backcolor = RGB(lc(0) + buff(0) * pro, lc(1) + buff(1) * pro, lc(2) + buff(2) * pro)
    ElseIf di = 2 Then
        di = 0: cd = False: X = ReleaseCapture
        UserControl.Backcolor = RGB(nc(0), nc(1), nc(2))
    Else
        Animation.Enabled = False
    End If
End Sub
Private Sub hover_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub hover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub hover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 0 And Y >= 0 And X <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight Then
        If cd = False Then
            cd = True: SetCapture UserControl.Hwnd
            inTime = GetTickCount: CopyMemory lc(0), UserControl.Backcolor, 4: CopyMemory nc(0), nhoc, 4
            Animation.Enabled = True
            di = 1
        End If
        RaiseEvent hover(Button, Shift, X, Y)
    Else
        If di = 1 Then
            di = 2
            inTime = GetTickCount: CopyMemory lc(0), UserControl.Backcolor, 4: CopyMemory nc(0), ndec, 4
            Animation.Enabled = True
            RaiseEvent Leave
        End If
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    RaiseEvent Click
    di = 0: cd = False: X = ReleaseCapture
    UserControl.Backcolor = ndec
End Sub
Private Sub UserControl_Resize()
    Select Case align
        Case 0
            con.Move UserControl.ScaleWidth / 2 - con.Width / 2, UserControl.ScaleHeight / 2 - con.Height / 2
        Case 1
            con.Move UserControl.ScaleHeight / 2 - con.Height / 2, UserControl.ScaleHeight / 2 - con.Height / 2
        Case 2
            con.Move UserControl.ScaleWidth - con.Width - (UserControl.ScaleHeight / 2 - con.Height / 2), UserControl.ScaleHeight / 2 - con.Height / 2
    End Select
    hover.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub
Private Sub UserControl_InitProperties()
    ndec = RGB(255, 255, 255)
    nhoc = RGB(242, 242, 242)
    UserControl.Backcolor = ndec
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ndec = PropBag.ReadProperty("DefaultColor", RGB(255, 255, 255))
    nhoc = PropBag.ReadProperty("HoverColor", RGB(242, 242, 242))
    na = PropBag.ReadProperty("Align", 0)
    con.ForeColor = PropBag.ReadProperty("ForeColor", RGB(64, 64, 64))
    Set con.Font = PropBag.ReadProperty("Font", con.Font)
    con.Caption = PropBag.ReadProperty("Content", "...")
    UserControl.Backcolor = ndec
    Call UserControl_Resize
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "DefaultColor", ndec, RGB(255, 255, 255)
    PropBag.WriteProperty "HoverColor", nhoc, RGB(242, 242, 242)
    PropBag.WriteProperty "ForeColor", con.ForeColor
    PropBag.WriteProperty "Font", con.Font
    PropBag.WriteProperty "Content", con.Caption
    PropBag.WriteProperty "Align", na
End Sub

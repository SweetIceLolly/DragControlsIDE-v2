VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Image.ocx"
Begin VB.Form frmPopupMenu 
   BackColor       =   &H001C1B1B&
   BorderStyle     =   0  'None
   Caption         =   "Dark¡áMenu"
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPopupTimeout 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   2040
   End
   Begin VB.Timer tmrCheckFocus 
      Interval        =   10
      Left            =   480
      Top             =   2040
   End
   Begin ImageX.aicAlphaImage imgMenuCheckBox 
      Height          =   345
      Index           =   0
      Left            =   1320
      Top             =   1080
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      Image           =   "frmPopupMenu.frx":0000
      Enabled         =   0   'False
   End
   Begin VB.Line lnSplitter 
      BorderColor     =   &H00373333&
      Index           =   0
      Visible         =   0   'False
      X1              =   600
      X2              =   1920
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image imgUnchecked 
      Enabled         =   0   'False
      Height          =   225
      Left            =   1320
      Picture         =   "frmPopupMenu.frx":0018
      Top             =   1560
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgChecked 
      Enabled         =   0   'False
      Height          =   225
      Left            =   1680
      Picture         =   "frmPopupMenu.frx":036E
      Top             =   1560
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgShowSubMenu 
      Enabled         =   0   'False
      Height          =   225
      Index           =   0
      Left            =   1320
      Picture         =   "frmPopupMenu.frx":06C4
      Top             =   720
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Line lnBorderTop 
      BorderColor     =   &H00373333&
      BorderWidth     =   2
      X1              =   840
      X2              =   120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line lnBorderLeft 
      BorderColor     =   &H00373333&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   360
      Y2              =   2040
   End
   Begin VB.Line lnBorderRight 
      BorderColor     =   &H00373333&
      BorderWidth     =   2
      X1              =   960
      X2              =   960
      Y1              =   600
      Y2              =   1680
   End
   Begin VB.Line lnBorderBottom 
      BorderColor     =   &H00373333&
      BorderWidth     =   2
      X1              =   240
      X2              =   1080
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label labItem 
      AutoSize        =   -1  'True
      BackColor       =   &H001C1B1B&
      Caption         =   " Item"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   225
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   525
   End
End
Attribute VB_Name = "frmPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ITEM_DISTANCE = 90
Private Const ITEM_HORZ_MARGIN = 60
Private Const SPLITTER_VERT_MARGIN = 90
Private Const MIN_WIDTH = 1425
Private Const TRANSPARENT_VALUE = 210

Private Type MenuItem
    MenuID          As Integer
    MenuText        As String
    SubMenus()      As String           'Base = 1
    SubMenuID()     As Integer          'Base = 1
    Enabled         As Boolean
    CheckBox        As Boolean
    Visible         As Boolean
    Checked         As Boolean
    MenuIcon()      As Byte
End Type

Dim Menus()         As MenuItem         'Base = 1
Dim CurrSubMenuID() As Integer          'Base = 1
Dim SpaceCount      As Integer
Dim SubMenuWindow   As frmPopupMenu
Dim BoundCtl        As DarkMenu
Dim LabelWidth      As Single
Dim PrevItem        As Integer
Public MatchItem    As Integer
Public IsPopupSub   As Boolean

Dim IsUsingKeyboard As Boolean
Dim KeybdIndex      As Integer
Public IsLastMenu   As Boolean
Dim PrevX           As Single, _
    PrevY           As Single

Public Sub CloseMenu()
    On Error Resume Next
    
    If Not SubMenuWindow Is Nothing Then
        SubMenuWindow.CloseMenu
        Set SubMenuWindow = Nothing
        IsPopupSub = False
    End If
    Unload Me
End Sub

Private Sub PopupNewMenu(LabelIndex As Integer)
    If Not SubMenuWindow Is Nothing Then
        SubMenuWindow.CloseMenu
        Set SubMenuWindow = Nothing
        IsPopupSub = False
    End If
    Set SubMenuWindow = New frmPopupMenu
    Me.IsPopupSub = True
    With SubMenuWindow
        .MatchItem = LabelIndex
        .Left = Me.Left + Me.labItem(LabelIndex).Width - 15
        .Top = Me.Top + Me.labItem(LabelIndex).Top - ITEM_DISTANCE
        .AddItems BoundCtl, Menus(CurrSubMenuID(LabelIndex + 1)).SubMenuID
        .Show
        If IsUsingKeyboard Then
            Call .Form_KeyDown(vbKeyDown, 0)
        End If
    End With
    If BoundCtl.Transparent Then
        SetLayeredWindowAttributes Me.hWnd, 0, TRANSPARENT_VALUE, LWA_ALPHA
    End If
End Sub

Public Sub AddItems(FromControl As DarkMenu, FromArray() As Integer, Optional ControlWidth As Integer)
    On Error Resume Next
    
    Dim i           As Integer
    Dim NewWidth    As Integer
    Dim HasCheckBox As Boolean
    Dim HasSubMenu  As Boolean
    Dim nCheckBoxes As Integer
    Dim nSubMenus   As Integer
    Dim nSplitters  As Integer
    
    CurrSubMenuID = FromArray
    SpaceCount = FromControl.SpaceCount
    Set BoundCtl = FromControl
    PrevItem = -1
    If BoundCtl.Transparent Then
        SetWindowLongA Me.hWnd, GWL_EXSTYLE, GetWindowLongA(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hWnd, 0, 255, LWA_ALPHA
    End If
    
    ReDim Menus(FromControl.GetMenuCount)
    For i = 1 To UBound(Menus)
        With Menus(i)
            FromControl.GetMenuItemInfo i, .MenuID, .MenuText, .Enabled, .CheckBox, _
                .Visible, .SubMenus, .SubMenuID, .MenuIcon, .Checked
        End With
    Next i
    
    For i = 1 To Me.labItem.ubound
        Unload Me.labItem(i)
    Next i
    
    HasCheckBox = False
    For i = 1 To UBound(CurrSubMenuID)
        If Menus(CurrSubMenuID(i)).Visible Then
            If i > 1 Then
                Load Me.labItem(i - 1)
            Else
                Me.labItem(0).AutoSize = True
                Me.labItem(0).Top = ITEM_DISTANCE
            End If
            Me.labItem(i - 1).Caption = String(SpaceCount, " ") & Menus(CurrSubMenuID(i)).MenuText & String(SpaceCount, " ")
            If i > 1 Then
                Me.labItem(i - 1).Top = Me.labItem(i - 2).Top + Me.labItem(i - 2).Height + ITEM_DISTANCE
                Me.labItem(i - 2).Height = Me.labItem(i - 1).Top - Me.labItem(i - 2).Top
            End If
            If Menus(CurrSubMenuID(i)).CheckBox = True And Menus(CurrSubMenuID(i)).MenuText <> "-" Then
                HasCheckBox = True
                If nCheckBoxes > 0 Then
                    Load Me.imgMenuCheckBox(nCheckBoxes)
                End If
                Me.imgMenuCheckBox(nCheckBoxes).Left = 60 'ITEM_HORZ_MARGIN
                Me.imgMenuCheckBox(nCheckBoxes).Top = Me.labItem(i - 1).Top + Me.labItem(i - 1).Height / 2 - Me.imgMenuCheckBox(nCheckBoxes).Height / 2
                If Menus(CurrSubMenuID(i)).Checked Then
                    Me.imgMenuCheckBox(nCheckBoxes).LoadImage_FromStdPicture Me.imgChecked.Picture
                Else
                    Me.imgMenuCheckBox(nCheckBoxes).LoadImage_FromStdPicture Me.imgUnchecked.Picture
                End If
                Me.imgMenuCheckBox(nCheckBoxes).Visible = True
                Me.imgMenuCheckBox(nCheckBoxes).ZOrder 0
                nCheckBoxes = nCheckBoxes + 1
            ElseIf (Not Menus(CurrSubMenuID(i)).MenuIcon) <> -1 And Menus(CurrSubMenuID(i)).MenuText <> "-" Then
                HasCheckBox = True
                If nCheckBoxes > 0 Then
                    Load Me.imgMenuCheckBox(nCheckBoxes)
                End If
                Me.imgMenuCheckBox(nCheckBoxes).Left = ITEM_HORZ_MARGIN + 90
                Me.imgMenuCheckBox(nCheckBoxes).Top = Me.labItem(i - 1).Top + Me.labItem(i - 1).Height / 2 - Me.imgMenuCheckBox(nCheckBoxes).Height / 2
                Me.imgMenuCheckBox(nCheckBoxes).LoadImage_FromArray Menus(CurrSubMenuID(i)).MenuIcon
                Me.imgMenuCheckBox(nCheckBoxes).Visible = True
                Me.imgMenuCheckBox(nCheckBoxes).ZOrder 0
                nCheckBoxes = nCheckBoxes + 1
            End If
            If UBound(Menus(CurrSubMenuID(i)).SubMenuID) > 0 Then
                HasSubMenu = True
                If nSubMenus > 0 Then
                    Load Me.imgShowSubMenu(nSubMenus)
                End If
                Me.imgShowSubMenu(nSubMenus).Top = Me.labItem(i - 1).Top + Me.labItem(i - 1).Height / 2 - Me.imgShowSubMenu(nSubMenus).Height / 2
                Me.imgShowSubMenu(nSubMenus).Visible = True
                nSubMenus = nSubMenus + 1
            End If
            Me.labItem(i - 1).Enabled = Menus(CurrSubMenuID(i)).Enabled
            If Me.labItem(i - 1).Width > NewWidth Then
                NewWidth = Me.labItem(i - 1).Width
            End If
            Me.labItem(i - 1).Visible = True
        Else
            If i > 1 Then
                Load Me.labItem(i - 1)
                Me.labItem(i - 1).Top = Me.labItem(i - 2).Top
                Me.labItem(i - 1).Height = Me.labItem(i - 2).Height
                Me.labItem(i - 1).Visible = False
            Else
                Me.labItem(0).Top = -Me.labItem(0).Height
                Me.labItem(0).Visible = False
            End If
        End If
    Next i
    
    If HasCheckBox Then
        For i = 0 To Me.labItem.ubound
            Me.labItem(i).Caption = "   " & Me.labItem(i).Caption
            If i > 0 Then
                Dim NextVisibleItem As Integer
                
                NextVisibleItem = i
                Do While Me.labItem(NextVisibleItem).Visible = False
                    NextVisibleItem = NextVisibleItem + 1
                    If NextVisibleItem = UBound(CurrSubMenuID) Then
                        Exit Do
                    End If
                Loop
                If NextVisibleItem < UBound(CurrSubMenuID) Then
                    Me.labItem(i - 1).Height = Me.labItem(NextVisibleItem).Top - Me.labItem(i - 1).Top
                End If
            End If
        Next i
        NewWidth = NewWidth + Me.imgMenuCheckBox(0).Width + ITEM_HORZ_MARGIN
    End If
    LabelWidth = ControlWidth
    Me.Height = Me.labItem(Me.labItem.ubound).Top + Me.labItem(Me.labItem.ubound).Height + ITEM_DISTANCE * 2
    For i = Me.labItem.ubound To 0 Step -1
        If Me.labItem(i).Visible = True Then
            Exit For
        End If
    Next i
    If i <> -1 Then
        Me.labItem(i).Height = Me.Height - Me.labItem(i).Top - ITEM_DISTANCE
        Me.Width = NewWidth + ITEM_HORZ_MARGIN * 2
        If Me.Width < MIN_WIDTH Then
            Me.Width = MIN_WIDTH
        End If
        For i = 0 To Me.labItem.ubound
            Me.labItem(i).Width = Me.Width
        Next i
        If HasSubMenu Then
            NewWidth = NewWidth + Me.imgShowSubMenu(0).Width + ITEM_HORZ_MARGIN
            For i = 0 To Me.imgShowSubMenu.ubound
                Me.imgShowSubMenu(i).Left = Me.Width - Me.imgShowSubMenu(0).Width - ITEM_HORZ_MARGIN * 2
                Me.imgShowSubMenu(i).ZOrder 0
            Next i
        End If
        For i = 0 To Me.labItem.ubound
            If Trim(Me.labItem(i).Caption) = "-" Then
                Me.labItem(i).Visible = False
                If nSplitters > 0 Then
                    Load Me.lnSplitter(nSplitters)
                End If
                With Me.lnSplitter(nSplitters)
                    .X1 = ITEM_HORZ_MARGIN * 2
                    .Y1 = Me.labItem(i).Top + SPLITTER_VERT_MARGIN
                    .X2 = Me.Width - ITEM_HORZ_MARGIN * 2
                    .Y2 = Me.labItem(i).Top + SPLITTER_VERT_MARGIN
                    .Visible = True
                End With
                nSplitters = nSplitters + 1
                
                '------------------------------------------------
                Dim j As Integer
                
                For j = i + 1 To Me.labItem.ubound
                    Me.labItem(j).Top = Me.labItem(j).Top - Me.labItem(i).Height + SPLITTER_VERT_MARGIN * 2
                Next j
                For j = 0 To Me.imgMenuCheckBox.ubound
                    If Me.imgMenuCheckBox(j).Top > Me.labItem(i).Top Then
                        Me.imgMenuCheckBox(j).Top = Me.imgMenuCheckBox(j).Top - Me.labItem(i).Height + SPLITTER_VERT_MARGIN * 2
                    End If
                Next j
                For j = 0 To Me.imgShowSubMenu.ubound
                    If Me.imgShowSubMenu(j).Top > Me.labItem(i).Top Then
                        Me.imgShowSubMenu(j).Top = Me.imgShowSubMenu(j).Top - Me.labItem(i).Height + SPLITTER_VERT_MARGIN * 2
                    End If
                Next j
                Me.Height = Me.Height - Me.labItem(i).Height + SPLITTER_VERT_MARGIN * 2
            End If
        Next i
    End If
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim PrevKeybdIndex      As Integer
    
    Select Case KeyCode
        Case vbKeyDown
            IsUsingKeyboard = True
            KeybdIndex = KeybdIndex + 1
            If KeybdIndex > Me.labItem.ubound Then
                KeybdIndex = 0
            End If
            PrevKeybdIndex = KeybdIndex
            Do While Menus(CurrSubMenuID(KeybdIndex + 1)).MenuText = "-" Or Menus(CurrSubMenuID(KeybdIndex + 1)).Enabled = False
                KeybdIndex = KeybdIndex + 1
                If KeybdIndex > Me.labItem.ubound Then
                    KeybdIndex = 0
                End If
                If KeybdIndex = PrevKeybdIndex Then
                    Exit Sub
                End If
            Loop
            Call labItem_MouseMove(KeybdIndex, 0, 0, 0, 0)
            If UBound(Menus(CurrSubMenuID(KeybdIndex + 1)).SubMenuID) = 0 Then
                Me.tmrPopupTimeout.Enabled = False
            End If
            IsUsingKeyboard = True
        
        Case vbKeyUp
            IsUsingKeyboard = True
            KeybdIndex = KeybdIndex - 1
            If KeybdIndex < 0 Then
                KeybdIndex = Me.labItem.ubound
            End If
            PrevKeybdIndex = KeybdIndex
            Do While Menus(CurrSubMenuID(KeybdIndex + 1)).MenuText = "-" Or Menus(CurrSubMenuID(KeybdIndex + 1)).Enabled = False
                KeybdIndex = KeybdIndex - 1
                If KeybdIndex < 0 Then
                    KeybdIndex = Me.labItem.ubound
                End If
                If KeybdIndex = PrevKeybdIndex Then
                    Exit Sub
                End If
            Loop
            Call labItem_MouseMove(KeybdIndex, 0, 0, 0, 0)
            If UBound(Menus(CurrSubMenuID(KeybdIndex + 1)).SubMenuID) = 0 Then
                Me.tmrPopupTimeout.Enabled = False
            End If
            IsUsingKeyboard = True
        
        Case vbKeyLeft
            IsUsingKeyboard = True
            Me.CloseMenu
            If IsLastMenu Then
                Call BoundCtl.MoveLeft
            End If
            IsUsingKeyboard = True
        
        Case vbKeyRight
            IsUsingKeyboard = True
            If Menus(CurrSubMenuID(KeybdIndex + 1)).Enabled Then
                If KeybdIndex = -1 Then
                    KeybdIndex = 0
                End If
                If UBound(Menus(CurrSubMenuID(KeybdIndex + 1)).SubMenuID) > 0 Then
                    Call labItem_MouseDown(KeybdIndex, 1, 0, 0, 0)
                Else
                    Call BoundCtl.MoveRight
                End If
            ElseIf KeybdIndex = -1 Then
                BoundCtl.MoveRight
            End If
            IsUsingKeyboard = True
        
        Case vbKeyReturn
            IsUsingKeyboard = True
            If Menus(CurrSubMenuID(KeybdIndex + 1)).Enabled Then
                Call labItem_MouseUp(KeybdIndex, vbLeftButton, 0, 0, 0)
            End If
            IsUsingKeyboard = True
        
        Case vbKeyEscape
            BoundCtl.HideMenu True
        
    End Select
End Sub

Private Sub Form_Load()
    KeybdIndex = -1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IsUsingKeyboard = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With Me.lnBorderBottom
        .X1 = 0
        .Y1 = Me.ScaleHeight - 15
        .X2 = Me.ScaleWidth
        .Y2 = Me.ScaleHeight - 15
    End With
    With Me.lnBorderLeft
        .X1 = 0
        .Y1 = 0
        .X2 = 0
        .Y2 = Me.ScaleHeight
    End With
    With Me.lnBorderTop
        .X1 = LabelWidth
        .Y1 = 0
        .X2 = Me.ScaleWidth
        .Y2 = 0
    End With
    With Me.lnBorderRight
        .X1 = Me.ScaleWidth - 15
        .Y1 = 0
        .X2 = Me.ScaleWidth - 15
        .Y2 = Me.ScaleHeight - 15
    End With
End Sub

Private Sub labItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        If UBound(Menus(CurrSubMenuID(Index + 1)).SubMenuID) > 0 Then
            If Not SubMenuWindow Is Nothing Then
                If SubMenuWindow.MatchItem <> Index Or SubMenuWindow.Visible = False Then
                    PopupNewMenu Index
                End If
            Else
                PopupNewMenu Index
            End If
            Me.tmrPopupTimeout.Enabled = False
        ElseIf Not SubMenuWindow Is Nothing Then
            If SubMenuWindow.MatchItem <> Index Then
                SubMenuWindow.CloseMenu
                Set SubMenuWindow = Nothing
            End If
        End If
    End If
End Sub

Private Sub labItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i           As Integer
    
    If Abs(PrevX - X) > 1 Or Abs(PrevY - Y) > 1 Or IsUsingKeyboard Then
        IsUsingKeyboard = False
        PrevX = X
        PrevY = Y
        If Index <> PrevItem Then
            For i = 0 To Me.labItem.ubound
                Me.labItem(i).BackColor = RGB(27, 27, 28)
            Next i
            Me.labItem(Index).BackColor = RGB(51, 51, 52)
            PrevItem = Index
            KeybdIndex = Index
            If UBound(Menus(CurrSubMenuID(Index + 1)).SubMenuID) > 0 Then
                Me.tmrPopupTimeout.Enabled = True
            Else
                Me.tmrPopupTimeout.Enabled = False
            End If
            If Not SubMenuWindow Is Nothing And Me.tmrPopupTimeout.Enabled = False Then
                If Index <> SubMenuWindow.MatchItem Then
                    Me.tmrPopupTimeout.Enabled = True
                Else
                    Me.tmrPopupTimeout.Enabled = False
                End If
            End If
        End If
    End If
End Sub

Private Sub labItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If UBound(Menus(CurrSubMenuID(Index + 1)).SubMenuID) = 0 Then
            BoundCtl.RaiseClickEvent Menus(CurrSubMenuID(Index + 1)).MenuID
            BoundCtl.HideMenu
            Unload Me
        End If
    End If
End Sub

Private Sub tmrCheckFocus_Timer()
    Dim pt          As POINT
    Dim i           As Integer
    
    GetCursorPos pt
    If WindowFromPoint(pt.X, pt.Y) <> Me.hWnd And Not IsUsingKeyboard Then
        PrevItem = -1
        Me.tmrPopupTimeout.Enabled = False
        For i = 0 To Me.labItem.ubound
            If Not SubMenuWindow Is Nothing Then
                If i <> SubMenuWindow.MatchItem Then
                    Me.labItem(i).BackColor = RGB(27, 27, 28)
                End If
            Else
                Me.labItem(i).BackColor = RGB(27, 27, 28)
            End If
        Next i
    End If
    If GetForegroundWindow <> Me.hWnd Then
        If SubMenuWindow Is Nothing Then
            Me.CloseMenu
        Else
            If GetForegroundWindow <> SubMenuWindow.hWnd And (Not SubMenuWindow.IsPopupSub) Then
                Me.CloseMenu
            End If
        End If
    End If
End Sub

Private Sub tmrPopupTimeout_Timer()
    If Not SubMenuWindow Is Nothing Then
        If PrevItem <> SubMenuWindow.MatchItem Then
            SetLayeredWindowAttributes Me.hWnd, 0, 255, LWA_ALPHA
            SubMenuWindow.CloseMenu
            Set SubMenuWindow = Nothing
            Me.tmrPopupTimeout.Enabled = False
            IsPopupSub = False
            If UBound(Menus(CurrSubMenuID(PrevItem + 1)).SubMenuID) > 0 Then
                PopupNewMenu PrevItem
            End If
            Exit Sub
        End If
    End If
    If SubMenuWindow Is Nothing Then
        PopupNewMenu PrevItem
    ElseIf SubMenuWindow.Visible = False Then
        PopupNewMenu PrevItem
    End If
    Me.tmrPopupTimeout.Enabled = False
End Sub

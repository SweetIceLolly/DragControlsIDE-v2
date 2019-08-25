VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Image.ocx"
Begin VB.UserControl TabBar 
   BackColor       =   &H00302D2D&
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   ScaleHeight     =   4485
   ScaleWidth      =   7425
   ToolboxBitmap   =   "TabBar.ctx":0000
   Begin VB.Timer DropInCheck 
      Interval        =   100
      Left            =   6912
      Top             =   528
   End
   Begin VB.Frame WindowFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00302D2D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3828
      Index           =   0
      Left            =   144
      TabIndex        =   7
      Top             =   528
      Visible         =   0   'False
      Width           =   7116
      Begin VB.Timer KeyCheckTimer 
         Interval        =   10
         Left            =   6768
         Top             =   960
      End
      Begin VB.Timer FixPosTimer 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6768
         Top             =   480
      End
   End
   Begin VB.Frame TopBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7476
      Begin VB.PictureBox MoreBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00302D2D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   7032
         ScaleHeight     =   420
         ScaleWidth      =   390
         TabIndex        =   9
         Top             =   0
         Width           =   396
         Begin VB.Image MoreBtnIcon 
            Height          =   165
            Left            =   150
            Picture         =   "TabBar.ctx":0312
            Top             =   150
            Width           =   240
         End
      End
      Begin VB.Label DropInMark 
         Alignment       =   2  'Center
         BackColor       =   &H003954FE&
         Height          =   420
         Left            =   6048
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1428
      End
      Begin VB.Label ClickCover 
         BackStyle       =   0  'Transparent
         Height          =   372
         Left            =   2448
         TabIndex        =   6
         Top             =   0
         Width           =   1524
      End
      Begin VB.Label CloseButton 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00CC7A00&
         Caption         =   "x"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Top             =   96
         Visible         =   0   'False
         Width           =   192
      End
      Begin VB.Label TabTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00CC7A00&
         BackStyle       =   0  'Transparent
         Caption         =   "test"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   432
         TabIndex        =   4
         Top             =   72
         Visible         =   0   'False
         Width           =   336
      End
      Begin ImageX.aicAlphaImage TabIcon 
         Height          =   216
         Index           =   0
         Left            =   96
         Top             =   96
         Visible         =   0   'False
         Width           =   252
         _ExtentX        =   450
         _ExtentY        =   370
         Image           =   "TabBar.ctx":0564
      End
      Begin VB.Label TabBg 
         BackColor       =   &H00CC7A00&
         Height          =   300
         Index           =   0
         Left            =   -1500
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label BottomBar 
         BackColor       =   &H00CC7A00&
         Height          =   24
         Left            =   24
         TabIndex        =   2
         Top             =   408
         Visible         =   0   'False
         Width           =   7428
      End
      Begin VB.Label BgCover 
         BackColor       =   &H00302D2D&
         Height          =   420
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7404
      End
   End
   Begin VB.Menu TabCollection 
      Caption         =   "标签页集合"
      Begin VB.Menu TabItem 
         Caption         =   "标签"
         Index           =   0
      End
   End
   Begin VB.Menu TabOperation 
      Caption         =   "标签操作"
      Begin VB.Menu cmdMoveOut 
         Caption         =   "切出"
      End
      Begin VB.Menu cmdClose 
         Caption         =   "关闭"
      End
   End
End
Attribute VB_Name = "TabBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================================
'描述:      标题栏控件
'作者:      Error 404, 冰棍
'文件:      TabBar.ctl
'====================================================

Event WindowDropIn(Frm As Form, Index As Integer)
Event WindowDropOut(Frm As Form, Index As Integer)
Event TabClick(Frm As Form, Index As Integer)

Dim FocusIndex As Integer, LastMove As Integer
Dim Windows() As Form
Dim SrcX  As Long, SrcY As Long, DropIndex As Integer, DropMode As Long
Dim SrcX2 As Long, SrcY2 As Long
Dim KeyClose As Boolean
Dim MenuTab As Integer
'NOTE：DeleteSource - 顺便删除原窗口，默认为True。
Public Sub RemoveFormByForm(Frm As Form, Optional DeleteSource As Boolean = True)
    Dim i As Integer
    For i = 1 To UBound(Windows)
        If Windows(i) Is Frm Then
            Call RemoveForm(i, DeleteSource)
            Exit For
        End If
    Next
End Sub
Public Sub RemoveForm(Index As Integer, Optional DeleteSource As Boolean = True)
    If Windows(Index) Is Nothing Then
        Exit Sub
    End If
    
    SetParent Windows(Index).hWnd, 0                            '在移除窗口前先恢复其母窗体，因为即使DeleteSource = False，窗口也会随着母窗口一起关闭
    If DeleteSource Then Unload Windows(Index)
    FixPosition Index, UBound(Windows) - 1, 1
    ReDim Preserve Windows(UBound(Windows) - 1)
    Unload TabIcon(TabIcon.UBound)
    Unload TabBg(TabBg.UBound)
    Unload TabTitle(TabTitle.UBound)
    Unload CloseButton(CloseButton.UBound)
    Unload WindowFrame(WindowFrame.UBound)
    If UBound(Windows) = 0 Then
        BottomBar.Visible = False
        LastMove = 0: FocusIndex = 0
    ElseIf UBound(Windows) >= Index Then
        LastMove = 0
        If FocusIndex > UBound(Windows) Then FocusIndex = Index
        If Index > UBound(Windows) Then
            SwitchTo UBound(Windows)
        Else
            SwitchTo Index
        End If
    Else
        LastMove = 0
        If FocusIndex > UBound(Windows) Then FocusIndex = Index - 1
        If Index - 1 > UBound(Windows) Then
            SwitchTo UBound(Windows)
        Else
            SwitchTo Index - 1
        End If
    End If
End Sub

Public Sub AddForm(Frm As Form)
    On Error Resume Next
    
    Dim Index As Integer
    Index = TabBg.UBound + 1
    
    Load TabTitle(Index)
    With TabTitle(Index)
        .Left = TabBg(Index - 1).Left + TabBg(Index - 1).Width + 16 * Screen.TwipsPerPixelX + TabIcon(Index - 1).Height
        .Top = (TopBar.Height - BottomBar.Height) / 2 - TabTitle(Index - 1).Height / 2
        .Caption = Frm.Caption
        .Tag = 16 * Screen.TwipsPerPixelX + TabIcon(Index - 1).Height
        .Visible = True
    End With
    
    Load TabBg(Index)
    With TabBg(Index)
        .Left = TabBg(Index - 1).Left + TabBg(Index - 1).Width
        .Top = 0
        .Width = TabTitle(Index).Width + TabTitle(Index).Height * 2 + 38 * Screen.TwipsPerPixelX
        .Height = TopBar.Height - BottomBar.Height
        .Visible = True
        .ZOrder 1
    End With
    
    Load TabIcon(Index)
    With TabIcon(Index)
        .Left = TabBg(Index - 1).Left + TabBg(Index - 1).Width + 8 * Screen.TwipsPerPixelX
        .Top = TabTitle(Index).Top
        .Visible = True
        .ClearImage
        .LoadImage_FromHandle SendMessageA(Frm.hWnd, WM_GETICON, ICON_BIG, 0)
        .Refresh
        .Width = TabTitle(Index).Height
        .Height = TabTitle(Index).Height
        .Visible = True
        .ZOrder
        .Tag = .Left - TabBg(Index).Left
    End With
    
    Load CloseButton(Index)
    With CloseButton(Index)
        .Left = TabBg(Index - 1).Left + TabBg(Index - 1).Width + 30 * Screen.TwipsPerPixelX + TabTitle(Index).Width + TabIcon(Index).Width
        .Top = TabTitle(Index).Top
        .Width = TabTitle(Index).Height
        .Height = TabTitle(Index).Height
        .Caption = "x"
        .Visible = True
        .Tag = .Left - TabBg(Index).Left
    End With
    
    BgCover.ZOrder 1
    CloseButton(Index).ZOrder
    
    BottomBar.Visible = True
    If Frm.Visible = False Then Frm.Show
    
    Frm.DarkTitleBar.Visible = False
    Frm.DarkWindowBorder.Bind = False
    Frm.DarkWindowBorderSizer.Bind = False
    Call Frm.Form_Resize
    
    Load WindowFrame(Index)
    With WindowFrame(Index)
        .Left = 0
        .Top = TopBar.Height
        .Visible = True
        .Tag = Frm.hWnd
        .ZOrder
        SetParent Frm.hWnd, .hWnd
    End With
    
    ReDim Preserve Windows(UBound(Windows) + 1)
    Set Windows(Index) = Frm
    
    SwitchTo Index
    
    RaiseEvent WindowDropIn(Windows(Index), Index)
End Sub

Public Sub SwitchTo(Index As Integer)
    TabBg(FocusIndex).BackColor = UserControl.BackColor
    TabBg(Index).BackColor = RGB(0, 122, 204)
    CloseButton(FocusIndex).BackColor = RGB(45, 45, 48)
    CloseButton(Index).BackColor = RGB(0, 122, 204)
    WindowFrame(FocusIndex).Visible = False
    WindowFrame(Index).Visible = True
    WindowFrame(Index).Width = UserControl.Width
    WindowFrame(Index).Height = UserControl.Height - TopBar.Height
    Windows(Index).Move 0, 0, UserControl.Width, UserControl.Height - TopBar.Height
    
    FocusIndex = Index
    
    RaiseEvent TabClick(Windows(Index), Index)
End Sub

'描述:      切换到指定的窗口
'参数:      TargetForm: 指定窗口
Public Sub SwitchToByForm(TargetForm As Form)
    Dim i   As Integer
    For i = 1 To UBound(Windows)                                        '尝试在窗口列表里找对应的窗口
        If Windows(i) Is TargetForm Then                                    '如果找到了就切换过去
            Call SwitchTo(i)
            Exit Sub
        End If
    Next i
    Call AddForm(TargetForm)                                            '找不到的话就先添加对应窗口，再切换过去
End Sub

'描述:      更新所有Tab的标题
Public Sub UpdateCaptions()
    Dim i   As Integer
    For i = 1 To UBound(Windows)
        TabTitle(i).Caption = Windows(i).Caption
    Next i
    FixPosition2 1, UBound(Windows)
End Sub

Private Sub ClickCover_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    DropIndex = 0
    For i = 1 To TabBg.UBound
        If X >= TabBg(i).Left And X <= TabBg(i).Left + TabBg(i).Width Then
            DropIndex = i: Exit For
        End If
    Next
    SrcX = X: SrcY = Y: DropMode = 0
    If DropIndex <> 0 Then SrcX2 = X - TabBg(DropIndex).Left: SrcY2 = Y - TabBg(DropIndex).Top
End Sub

Private Sub ClickCover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If LastMove <> 0 Then
        CloseButton(LastMove).BackColor = IIf(FocusIndex = LastMove, RGB(0, 122, 204), RGB(45, 45, 48))
        LastMove = 0
    End If
    If Button = 1 And DropIndex <> 0 Then
        If Abs(X - SrcX) >= 8 * Screen.TwipsPerPixelX And DropMode = 0 Then
            DropMode = 1
            TabBg(DropIndex).ZOrder
            TabIcon(DropIndex).ZOrder
            TabTitle(DropIndex).ZOrder
            CloseButton(DropIndex).ZOrder
        End If
        If Abs(Y - SrcY) >= 8 * Screen.TwipsPerPixelY And DropMode = 0 Then
            DropMode = 2
            TabBg(DropIndex).BackColor = RGB(254, 84, 57)
            TabTitle(DropIndex).Caption = "*" & TabTitle(DropIndex)
            CloseButton(DropIndex).BackColor = RGB(254, 84, 57)
            BottomBar.BackColor = RGB(254, 84, 57)
            Dim p As POINTAPI
            GetCursorPos p
            SetWindowPos Windows(DropIndex).hWnd, 0, p.X - SrcX / Screen.TwipsPerPixelX, p.Y - SrcY / Screen.TwipsPerPixelY, 0, 0, SWP_NOZORDER Or SWP_NOSIZE
            Windows(DropIndex).DarkTitleBar.Visible = True
            Windows(DropIndex).DarkWindowBorder.Bind = True
            Windows(DropIndex).DarkWindowBorderSizer.Bind = True
            Call Windows(DropIndex).Form_Resize
            SetParent Windows(DropIndex).hWnd, 0
            Dim rtn As Long
            rtn = GetWindowLongA(Windows(DropIndex).hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
            SetWindowLongA Windows(DropIndex).hWnd, GWL_EXSTYLE, rtn
            SetLayeredWindowAttributes Windows(DropIndex).hWnd, 0, 128, LWA_ALPHA
            
            Sleep 100: DoEvents
            Do While DropMode = 2
                GetCursorPos p
                SetWindowPos Windows(DropIndex).hWnd, 0, p.X - SrcX / Screen.TwipsPerPixelX, p.Y - SrcY / Screen.TwipsPerPixelY, 0, 0, SWP_NOZORDER Or SWP_NOSIZE
                Sleep 10: DoEvents
            Loop
        End If
        If DropMode = 1 Then
            TabBg(DropIndex).Left = X - SrcX2
            TabIcon(DropIndex).Left = TabBg(DropIndex).Left + Val(TabIcon(DropIndex).Tag)
            TabTitle(DropIndex).Left = TabBg(DropIndex).Left + Val(TabTitle(DropIndex).Tag)
            CloseButton(DropIndex).Left = TabBg(DropIndex).Left + Val(CloseButton(DropIndex).Tag)
        End If
    End If
End Sub

Private Sub ClickCover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DropMode = 2 Then
        SetLayeredWindowAttributes Windows(DropIndex).hWnd, 0, 255, LWA_ALPHA
        RaiseEvent WindowDropOut(Windows(DropIndex), DropIndex)
        Dim rtn As Long
        rtn = GetWindowLongA(Windows(DropIndex).hWnd, GWL_EXSTYLE) And (Not WS_EX_LAYERED)
        SetWindowLongA Windows(DropIndex).hWnd, GWL_EXSTYLE, rtn
        RemoveForm DropIndex, False
        BottomBar.BackColor = RGB(0, 122, 204)
    End If
    
    If DropIndex <> 0 And Button = 1 Then
        'If FocusIndex <> DropIndex And DropMode = 0 Then SwitchTo DropIndex
        If DropMode = 0 Then SwitchTo DropIndex                             '在此处修改成如果DropMode=0就触发SwitchTo，修复文本框不获取焦点的问题
        If DropMode = 1 Then
            TabTitle(DropIndex).ZOrder 1
            TabBg(DropIndex).ZOrder 1
            BgCover.ZOrder 1
            Dim i As Integer, NewIndex As Integer
            NewIndex = DropIndex
            If X < 0 Then
                NewIndex = 1
            ElseIf X > UserControl.Width And TabBg(UBound(Windows)).Left + TabBg(UBound(Windows)).Width < UserControl.Width Then
                NewIndex = UBound(Windows)
            Else
                For i = 1 To TabBg.UBound
                    If X >= TabBg(i).Left And X <= TabBg(i).Left + TabBg(i).Width And i <> DropIndex Then NewIndex = i: Exit For
                Next
            End If
            
            If NewIndex <> DropIndex Then
                Dim tempW As Form
                Set tempW = Windows(DropIndex)
                If NewIndex < DropIndex Then
                    For i = DropIndex To NewIndex + 1 Step -1
                        Set Windows(i) = Windows(i - 1)
                        SetParent Windows(i).hWnd, WindowFrame(i).hWnd
                        TabTitle(i).Caption = Windows(i).Caption
                    Next
                Else
                    For i = DropIndex To NewIndex - 1
                        Set Windows(i) = Windows(i + 1)
                        SetParent Windows(i).hWnd, WindowFrame(i).hWnd
                        TabTitle(i).Caption = Windows(i).Caption
                    Next
                End If
                Set Windows(NewIndex) = tempW
                TabTitle(NewIndex).Caption = Windows(NewIndex).Caption
                SetParent Windows(NewIndex).hWnd, WindowFrame(NewIndex).hWnd
            End If
            If NewIndex < DropIndex Then
                FixPosition2 NewIndex, DropIndex
            Else
                FixPosition2 DropIndex, NewIndex
            End If
            SwitchTo NewIndex
        End If
    End If
    
    SrcX = 0: SrcY = 0: SrcX2 = 0: SrcY2 = 0
    
    DropMode = 0: DropIndex = 0
End Sub

Private Sub CloseButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    CloseButton(Index).BackColor = IIf(FocusIndex = Index, RGB(28, 151, 234), RGB(96, 96, 96))
    LastMove = Index
End Sub

Private Sub CloseButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        RemoveForm Index
    End If
End Sub

Private Sub FixPosition(StartPos As Integer, EndPos As Integer, Step As Integer)
    Dim cx As Long
    For i = StartPos To EndPos
        Set Windows(i) = Windows(i + Step)
        cx = TabBg(i + Step).Left - (TabBg(i - 1).Left + TabBg(i - 1).Width)
        TabIcon(i).Left = TabIcon(i + Step).Left - cx
        TabBg(i).Left = TabBg(i + Step).Left - cx
        TabTitle(i).Left = TabTitle(i + Step).Left - cx
        CloseButton(i).Left = CloseButton(i + Step).Left - cx
        TabTitle(i).Caption = TabTitle(i + Step).Caption
        TabBg(i).Width = TabBg(i + Step).Width
        TabIcon(i).LoadImage_FromHandle SendMessageA(Windows(i).hWnd, WM_GETICON, ICON_BIG, 0)
        SetParent Windows(i).hWnd, WindowFrame(i).hWnd
        CloseButton(i).Tag = CloseButton(i).Left - TabBg(i).Left
    Next
End Sub

Private Sub FixPosition2(StartPos As Integer, EndPos As Integer)
    Dim cx As Long
    For i = StartPos To EndPos
        TabBg(i).Left = TabBg(i - 1).Left + TabBg(i - 1).Width
        TabIcon(i).Left = TabBg(i).Left + Val(TabIcon(i).Tag)
        TabTitle(i).Left = TabBg(i).Left + Val(TabTitle(i).Tag)
        TabBg(i).Width = TabTitle(i).Width + TabTitle(i).Height * 2 + 38 * Screen.TwipsPerPixelX
        CloseButton(i).Left = TabBg(i).Left + 30 * Screen.TwipsPerPixelX + TabTitle(i).Width + TabIcon(i).Width
        CloseButton(i).Tag = CloseButton(i).Left - TabBg(i).Left
        TabIcon(i).LoadImage_FromHandle SendMessageA(Windows(i).hWnd, WM_GETICON, ICON_BIG, 0)
    Next
End Sub

Private Sub cmdClose_Click()
    If MenuTab = 0 Then Exit Sub
    RemoveForm MenuTab
    MenuTab = 0
End Sub

Private Sub cmdMoveOut_Click()
    If MenuTab = 0 Then Exit Sub
    Dim p As POINTAPI
    GetCursorPos p
    SetWindowPos Windows(MenuTab).hWnd, 0, p.X, p.Y, 0, 0, SWP_NOZORDER Or SWP_NOSIZE
    SetParent Windows(MenuTab).hWnd, 0
    RaiseEvent WindowDropOut(Windows(MenuTab), MenuTab)
    RemoveForm MenuTab, False
    MenuTab = 0
End Sub

Private Sub DropInCheck_Timer()
    If Not Ambient.UserMode Then
        UserControl.DropInCheck.Enabled = False
    End If
    
    If UserControl.Extender.Visible = False Then Exit Sub
    
    Dim hWnd As Long, MBtn As Boolean
    hWnd = GetActiveWindow
    If hWnd = 0 Then Exit Sub
    If DropMode = 2 Then Exit Sub
    
    MBtn = (GetAsyncKeyState(VK_LBUTTON) <> 0)
    
    If Not MBtn Then Exit Sub
    
    Dim r As RECT, p As POINTAPI
    GetWindowRect hWnd, r
    GetCursorPos p
    
    If Not (p.X >= r.Left And p.X <= r.Right And p.Y >= r.Top And p.Y - r.Top <= 33 * (15 / Screen.TwipsPerPixelY)) Then Exit Sub
    
    GetWindowRect UserControl.hWnd, r
    
    If Not (p.Y >= r.Top And p.Y <= r.Top + TopBar.Height / Screen.TwipsPerPixelY And p.X >= r.Left And p.X <= r.Right) Then Exit Sub
    
    Dim Frm As Form, DropForm As Form
    
    For Each Frm In VB.Forms
        If Frm.hWnd = hWnd Then Set DropForm = Frm: Exit For
    Next
    
    If DropForm Is Nothing Then Exit Sub
    On Error GoTo FailRead
    If DropForm.DarkTitleBar Is Nothing Then Exit Sub
    GoTo SuccessRead
    
FailRead:
    Exit Sub
SuccessRead:
    
    Dim X As Single
    
    BottomBar.BackColor = RGB(254, 84, 57)
    
    Dim rtn As Long
    rtn = GetWindowLongA(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetWindowLongA hWnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hWnd, 0, 100, LWA_ALPHA
    
    DropInMark.Visible = True
    DropInMark.ZOrder
    
    Dim Text As String * 255
    GetWindowTextA hWnd, Text, 255
    DropInMark.Caption = "*" & Text
    DropInMark.ForeColor = RGB(255, 255, 255)
    DropInMark.Width = Len(DropInMark.Caption) * DropInMark.FontSize * Screen.TwipsPerPixelX + 50 * Screen.TwipsPerPixelX
    
    Do While MBtn
        GetCursorPos p
        MBtn = (GetAsyncKeyState(VK_LBUTTON) <> 0)
        X = (p.X - r.Left) * Screen.TwipsPerPixelX
        DropInMark.Left = X - DropInMark.Width / 2
        Sleep 10: DoEvents
    Loop
    
    SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
    rtn = rtn And (Not WS_EX_LAYERED)
    SetWindowLongA hWnd, GWL_EXSTYLE, rtn
    
    BottomBar.BackColor = RGB(0, 122, 204)
    DropInMark.Visible = False
    
    If X < 0 Or X > UserControl.Width Then Exit Sub
    
    AddForm DropForm
    'Cheat
    DropMode = 1: DropIndex = UBound(Windows)
    Call ClickCover_MouseUp(1, 0, X, 0)
    
    UserControl.SetFocus
    FixPosTimer.Enabled = True
End Sub

Private Sub FixPosTimer_Timer()
    If Not Ambient.UserMode Then
        UserControl.FixPosTimer.Enabled = False
    End If
    
    SwitchTo FocusIndex
    FixPosTimer.Enabled = False
End Sub

Private Sub KeyCheckTimer_Timer()
    If Not Ambient.UserMode Then
        UserControl.KeyCheckTimer.Enabled = False
    End If
    
    If UserControl.Extender.Visible = False Then Exit Sub
    
    If GetActiveWindow = 0 Then Exit Sub

    If (GetAsyncKeyState(VK_CONTROL) <> 0) And (GetAsyncKeyState(VK_W) <> 0) Then
        If Not KeyClose Then
            KeyClose = True
            RemoveForm FocusIndex
        End If
    Else
        KeyClose = False
    End If
End Sub

Private Sub MoreBtn_Click()
    If UBound(Windows) = 0 Then Exit Sub
    
    UserControl.TabItem(0).Visible = True
    If UserControl.TabItem.UBound > 0 Then
        For i = 1 To UserControl.TabItem.UBound
            Unload UserControl.TabItem(i)
        Next
    End If
    
    For i = 1 To UBound(Windows)
        Load UserControl.TabItem(i)
        UserControl.TabItem(i).Caption = Windows(i).Caption
    Next
    UserControl.TabItem(0).Visible = False
    
    UserControl.PopupMenu TabCollection
End Sub

Private Sub MoreBtnIcon_Click()
    Call MoreBtn_Click
End Sub

Private Sub TabIcon_Click(Index As Integer, ByVal Button As Integer)
    If Button = 2 Then
        MenuTab = Index
        UserControl.PopupMenu TabOperation
    End If
End Sub

Private Sub TabItem_Click(Index As Integer)
    If Index > UBound(Windows) Then Exit Sub
    SwitchTo Index
End Sub

Private Sub UserControl_Initialize()
    ReDim Windows(0)
    TabBg(0).Left = -TabBg(0).Width
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    BottomBar.Height = 3 * Screen.TwipsPerPixelY
    BottomBar.Width = UserControl.Width
    TopBar.Width = UserControl.Width
    TopBar.Height = (TabTitle(0).Height + 16 * Screen.TwipsPerPixelX) + BottomBar.Height
    BottomBar.Top = TopBar.Height - BottomBar.Height
    BgCover.Move 0, 0, TopBar.Width, TopBar.Height
    ClickCover.Move 0, 0, TopBar.Width, TopBar.Height
    DropInMark.Height = TopBar.Height
    
    MoreBtn.Height = TopBar.Height - BottomBar.Height
    MoreBtn.Width = MoreBtn.Height
    MoreBtnIcon.Left = MoreBtn.Width / 2 - MoreBtnIcon.Width / 2
    MoreBtnIcon.Top = MoreBtn.Height / 2 - MoreBtnIcon.Height / 2
    MoreBtn.Left = UserControl.Width - MoreBtn.Width
    
    If FocusIndex = 0 Then Exit Sub
    WindowFrame(FocusIndex).Width = UserControl.Width
    WindowFrame(FocusIndex).Height = UserControl.Height - TopBar.Height
    Windows(FocusIndex).Move 0, 0, UserControl.Width, UserControl.Height - TopBar.Height
End Sub

VERSION 5.00
Begin VB.Form frmBorderWindow 
   BorderStyle     =   0  'None
   Caption         =   "Dark¡áBorder"
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1305
   Icon            =   "frmBorderWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   1305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCheckFocus 
      Interval        =   10
      Left            =   -120
      Top             =   120
   End
End
Attribute VB_Name = "frmBorderWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dark¡áBorderWindow by IceLolly
'Date: 2018.8.8

'Please note that this window is used by Dark¡áWindowBorder and is NOT suggested to use directly.

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function I_want_to_fuck_Visual_Basic_since_it_does_not_let_me_to_use_SetFocus_as_a_function_name Lib _
    "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

Enum BindPosition
    TopPos = 0
    LeftPos = 1
    RightPos = 2
    BottomPos = 3
End Enum

Private Const FACTOR = 5 * 15

Public BoundWindow  As Long
Public UseSetParent As Boolean
Public CanSize      As Boolean
Public Thickness    As Integer
Public fColor       As Long
Public nfColor      As Long
Public MinW         As Long
Public MinH         As Long

Dim Pos             As BindPosition
Dim Trans           As Byte
Dim SizeMode        As Integer
'0  1   2
' ¨I¡ü¨J
'3¡û  ¡ú4
' ¨L¡ý¨K
'5  6   7
Dim bMoving         As Boolean
Dim PrevRect        As RECT
Dim NewX            As Long, NewY       As Long, _
    NewW            As Long, NewH       As Long

Public Property Get BindPos() As BindPosition
    BindPos = Pos
End Property

Public Property Let BindPos(DesirePos As BindPosition)
    Pos = DesirePos
    
    Select Case DesirePos
        Case TopPos
            Me.MousePointer = 7
            
        Case BottomPos
            Me.MousePointer = 7
        
        Case LeftPos
            Me.MousePointer = 9
        
        Case RightPos
            Me.MousePointer = 9
    
    End Select
End Property

Public Property Get Transparency() As Byte
    Transparency = Trans
End Property

Public Property Let Transparency(NewValue As Byte)
    Trans = NewValue
    
    SetWindowLongA Me.hWnd, GWL_EXSTYLE, GetWindowLongA(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, 0, NewValue, LWA_ALPHA
End Property

Private Sub Form_GotFocus()
    I_want_to_fuck_Visual_Basic_since_it_does_not_let_me_to_use_SetFocus_as_a_function_name BoundWindow
End Sub

Private Sub Form_Load()
    SetWindowLongA Me.hWnd, GWL_STYLE, GetWindowLongA(Me.hWnd, GWL_STYLE) Or WS_CHILD
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If CanSize Then
        If Button = 1 Then
            Dim ParentHwnd  As Long
            Dim ParentRect  As RECT
            
            ParentHwnd = GetPropA(BoundWindow, "Parent")
            GetWindowRect BoundWindow, PrevRect
            If ParentHwnd <> 0 Then
                GetWindowRect ParentHwnd, ParentRect
                PrevRect.Left = PrevRect.Left - ParentRect.Left
                PrevRect.Top = PrevRect.Top - ParentRect.Top
                ClipCursor ParentRect
            End If
            bMoving = True
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMoving = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bMoving Then
        Exit Sub
    End If
    If Not CanSize Then
        Me.MousePointer = 0
        Exit Sub
    End If
    If Pos = TopPos Then
        If X < Thickness * FACTOR Then
            Me.MousePointer = 8
            SizeMode = 0
        ElseIf X > Me.Width - Thickness * FACTOR Then
            Me.MousePointer = 6
            SizeMode = 2
        Else
            Me.MousePointer = 7
            SizeMode = 1
        End If
    ElseIf Pos = BottomPos Then
        If X < Thickness * FACTOR Then
            Me.MousePointer = 6
            SizeMode = 5
        ElseIf X > Me.Width - Thickness * FACTOR Then
            Me.MousePointer = 8
            SizeMode = 7
        Else
            Me.MousePointer = 7
            SizeMode = 6
        End If
    ElseIf Pos = LeftPos Then
        If Y < Thickness * (FACTOR - 1) Then
            Me.MousePointer = 8
            SizeMode = 0
        ElseIf Y > Me.Height - Thickness * (FACTOR - 1) Then
            Me.MousePointer = 6
            SizeMode = 5
        Else
            Me.MousePointer = 9
            SizeMode = 3
        End If
    ElseIf Pos = RightPos Then
        If Y < Thickness * (FACTOR - 1) Then
            Me.MousePointer = 6
            SizeMode = 2
        ElseIf Y > Me.Height - Thickness * (FACTOR - 1) Then
            Me.MousePointer = 8
            SizeMode = 7
        Else
            Me.MousePointer = 9
            SizeMode = 4
        End If
    End If
End Sub

Private Sub tmrCheckFocus_Timer()
    If GetForegroundWindow = BoundWindow Then
        Me.BackColor = fColor
    Else
        Me.BackColor = nfColor
    End If
    
    '-------------------------------------------------
    Dim wp          As WINDOWPLACEMENT
    Dim wRect       As RECT
    
    GetWindowPlacement BoundWindow, wp
    If wp.ShowCmd = SW_MAXIMIZE Then
        Me.Hide
    Else
        GetWindowRect BoundWindow, wRect
        If Me.Visible = False Then
            Me.Show
        End If
        Select Case Pos
            Case TopPos
                If UseSetParent Then
                    SetWindowPos Me.hWnd, 0, 0, 0, _
                        wRect.Right - wRect.Left, Thickness, SWP_NOZORDER Or SWP_NOACTIVATE
                Else
                    SetWindowPos Me.hWnd, 0, wRect.Left - Thickness, wRect.Top - Thickness, _
                        wRect.Right - wRect.Left + Thickness * 2, Thickness, SWP_NOZORDER Or SWP_NOACTIVATE
                End If
            
            Case BottomPos
                If UseSetParent Then
                    SetWindowPos Me.hWnd, 0, 0, wRect.bottom - wRect.Top - Thickness, _
                        wRect.Right - wRect.Left, Thickness, SWP_NOZORDER Or SWP_NOACTIVATE
                Else
                    SetWindowPos Me.hWnd, 0, wRect.Left - Thickness, wRect.bottom, _
                        wRect.Right - wRect.Left + Thickness * 2, Thickness, SWP_NOZORDER Or SWP_NOACTIVATE
                End If
                
            Case LeftPos
                If UseSetParent Then
                    SetWindowPos Me.hWnd, 0, 0, 0, _
                        Thickness, wRect.bottom - wRect.Top, SWP_NOZORDER Or SWP_NOACTIVATE
                Else
                    SetWindowPos Me.hWnd, 0, wRect.Left - Thickness, wRect.Top, _
                        Thickness, wRect.bottom - wRect.Top, SWP_NOZORDER Or SWP_NOACTIVATE
                End If
                
            Case RightPos
                If UseSetParent Then
                    SetWindowPos Me.hWnd, 0, wRect.Right - wRect.Left - Thickness, 0, _
                        Thickness, wRect.bottom - wRect.Top, SWP_NOZORDER Or SWP_NOACTIVATE
                Else
                    SetWindowPos Me.hWnd, 0, wRect.Right, wRect.Top, _
                        Thickness, wRect.bottom - wRect.Top, SWP_NOZORDER Or SWP_NOACTIVATE
                End If
                
        End Select
        
        If GetAsyncKeyState(VK_LBUTTON) = 0 Then
            ClipCursor ByVal 0
            bMoving = False
        End If
        If bMoving Then
            Dim cur             As POINT
            Dim ParentWindow    As Long
            Dim ParentRect      As RECT
            
            GetCursorPos cur
            ParentWindow = GetPropA(BoundWindow, "Parent")
            If ParentWindow <> 0 Then
                GetWindowRect ParentWindow, ParentRect
            End If
            Select Case SizeMode
                Case 0
                    If ParentWindow = 0 Then
                        If cur.X < PrevRect.Right Then
                            NewX = cur.X
                            NewW = PrevRect.Right - cur.X
                        End If
                        If cur.Y < PrevRect.bottom Then
                            NewY = cur.Y
                            NewH = PrevRect.bottom - cur.Y
                        End If
                        If NewW < MinW Then
                            NewW = MinW
                            NewX = PrevRect.Right - NewW
                        End If
                        If NewH < MinH Then
                            NewH = MinH
                            NewY = PrevRect.bottom - NewH
                        End If
                    Else
                        If cur.X < PrevRect.Right Then
                            NewX = cur.X - ParentRect.Left
                            NewW = PrevRect.Right - cur.X
                        End If
                        If cur.Y < PrevRect.bottom Then
                            NewY = cur.Y - ParentRect.Top
                            NewH = PrevRect.bottom - cur.Y
                        End If
                        If NewW < MinW Then
                            NewW = MinW
                            NewX = PrevRect.Right - NewW - ParentRect.Left
                        End If
                        If NewH < MinH Then
                            NewH = MinH
                            NewY = PrevRect.bottom - NewH - ParentRect.Top
                        End If
                    End If
                
                Case 1
                    If ParentWindow = 0 Then
                        NewX = PrevRect.Left
                        NewW = PrevRect.Right - PrevRect.Left
                        If cur.Y < PrevRect.bottom Then
                            NewY = cur.Y
                            NewH = PrevRect.bottom - cur.Y
                        End If
                        If NewH < MinH Then
                            NewH = MinH
                            NewY = PrevRect.bottom - NewH
                        End If
                    Else
                        NewX = PrevRect.Left
                        NewW = PrevRect.Right - PrevRect.Left - ParentRect.Left
                        If cur.Y < PrevRect.bottom Then
                            NewY = cur.Y - ParentRect.Top
                            NewH = PrevRect.bottom - cur.Y
                        End If
                        If NewH < MinH Then
                            NewH = MinH
                            NewY = PrevRect.bottom - NewH - ParentRect.Top
                        End If
                    End If
                
                Case 2
                    If ParentWindow = 0 Then
                        NewX = PrevRect.Left
                        If cur.X > PrevRect.Left Then
                            NewW = cur.X - PrevRect.Left
                        End If
                        If cur.Y < PrevRect.bottom Then
                            NewY = cur.Y
                            NewH = PrevRect.bottom - cur.Y
                        End If
                        If NewW < MinW Then
                            NewW = MinW
                        End If
                        If NewH < MinH Then
                            NewH = MinH
                            NewY = PrevRect.bottom - NewH
                        End If
                    Else
                        NewX = PrevRect.Left
                        If cur.X > PrevRect.Left Then
                            NewW = cur.X - PrevRect.Left - ParentRect.Left
                        End If
                        If cur.Y < PrevRect.bottom Then
                            NewY = cur.Y - ParentRect.Top
                            NewH = PrevRect.bottom - cur.Y
                        End If
                        If NewW < MinW Then
                            NewW = MinW
                        End If
                        If NewH < MinH Then
                            NewH = MinH
                            NewY = PrevRect.bottom - NewH - ParentRect.Top
                        End If
                    End If
                
                Case 3
                    If ParentWindow = 0 Then
                        If cur.X < PrevRect.Right Then
                            NewX = cur.X
                            NewW = PrevRect.Right - cur.X
                        End If
                        NewY = PrevRect.Top
                        NewH = PrevRect.bottom - PrevRect.Top
                        If NewW < MinW Then
                            NewW = MinW
                            NewX = PrevRect.Right - NewW
                        End If
                    Else
                        If cur.X < PrevRect.Right Then
                            NewX = cur.X - ParentRect.Left
                            NewW = PrevRect.Right - cur.X
                        End If
                        NewY = PrevRect.Top
                        NewH = PrevRect.bottom - PrevRect.Top - ParentRect.Top
                        If NewW < MinW Then
                            NewW = MinW
                            NewX = PrevRect.Right - NewW - ParentRect.Left
                        End If
                    End If
                
                Case 4
                    If ParentWindow = 0 Then
                        NewX = PrevRect.Left
                        If cur.X > PrevRect.Left Then
                            NewW = cur.X - PrevRect.Left
                        End If
                        NewY = PrevRect.Top
                        NewH = PrevRect.bottom - PrevRect.Top
                        If NewW < MinW Then
                            NewW = MinW
                        End If
                    Else
                        NewX = PrevRect.Left
                        If cur.X > PrevRect.Left Then
                            NewW = cur.X - PrevRect.Left - ParentRect.Left
                        End If
                        NewY = PrevRect.Top
                        NewH = PrevRect.bottom - PrevRect.Top - ParentRect.Top
                        If NewW < MinW Then
                            NewW = MinW
                        End If
                    End If
                
                Case 5
                    If ParentWindow = 0 Then
                        If cur.X < PrevRect.Right Then
                            NewX = cur.X
                            NewW = PrevRect.Right - cur.X
                        End If
                        NewY = PrevRect.Top
                        If cur.Y > PrevRect.Top Then
                            NewH = cur.Y - PrevRect.Top
                        End If
                        If NewW < MinW Then
                            NewW = MinW
                            NewX = PrevRect.Right - NewW
                        End If
                        If NewH < MinH Then
                            NewH = MinH
                        End If
                    Else
                        If cur.X < PrevRect.Right Then
                            NewX = cur.X - ParentRect.Left
                            NewW = PrevRect.Right - cur.X
                        End If
                        NewY = PrevRect.Top
                        If cur.Y > PrevRect.Top Then
                            NewH = cur.Y - PrevRect.Top - ParentRect.Top
                        End If
                        If NewW < MinW Then
                            NewW = MinW
                            NewX = PrevRect.Right - NewW - ParentRect.Left
                        End If
                        If NewH < MinH Then
                            NewH = MinH
                        End If
                    End If
                
                Case 6
                    If ParentWindow = 0 Then
                        NewX = PrevRect.Left
                        NewW = PrevRect.Right - PrevRect.Left
                        NewY = PrevRect.Top
                        If cur.Y > PrevRect.Top Then
                            NewH = cur.Y - PrevRect.Top
                        End If
                        If NewH < MinH Then
                            NewH = MinH
                        End If
                    Else
                        NewX = PrevRect.Left
                        NewW = PrevRect.Right - PrevRect.Left - ParentRect.Left
                        NewY = PrevRect.Top
                        If cur.Y > PrevRect.Top Then
                            NewH = cur.Y - PrevRect.Top - ParentRect.Top
                        End If
                        If NewH < MinH Then
                            NewH = MinH
                        End If
                    End If
                
                Case 7
                    If ParentWindow = 0 Then
                        NewX = PrevRect.Left
                        If cur.X > PrevRect.Left Then
                            NewW = cur.X - PrevRect.Left
                        End If
                        NewY = PrevRect.Top
                        If cur.Y > PrevRect.Top Then
                            NewH = cur.Y - PrevRect.Top
                        End If
                        If NewW < MinW Then
                            NewW = MinW
                        End If
                        If NewH < MinH Then
                            NewH = MinH
                        End If
                    Else
                        NewX = PrevRect.Left
                        If cur.X > PrevRect.Left Then
                            NewW = cur.X - PrevRect.Left - ParentRect.Left
                        End If
                        NewY = PrevRect.Top
                        If cur.Y > PrevRect.Top Then
                            NewH = cur.Y - PrevRect.Top - ParentRect.Top
                        End If
                        If NewW < MinW Then
                            NewW = MinW
                        End If
                        If NewH < MinH Then
                            NewH = MinH
                        End If
                    End If
                
            End Select
            SetWindowPos BoundWindow, 0, NewX, NewY, NewW, NewH, SWP_NOZORDER
        End If
    End If
End Sub

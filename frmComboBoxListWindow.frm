VERSION 5.00
Begin VB.Form frmComboBoxListWindow 
   BackColor       =   &H001C1B1B&
   BorderStyle     =   0  'None
   Caption         =   "Dark¡áComboList"
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   Icon            =   "frmComboBoxListWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   1500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCheckFocus 
      Interval        =   10
      Left            =   360
      Top             =   1920
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H001C1B1B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   975
      TabIndex        =   1
      Top             =   0
      Width           =   975
      Begin VB.Label labItem 
         AutoSize        =   -1  'True
         BackColor       =   &H001C1B1B&
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F0F0F0&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin DragControlsIDE.DarkVScrollBar VscrollBar 
      Height          =   1935
      Left            =   1080
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3413
   End
End
Attribute VB_Name = "frmComboBoxListWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dark¡áComboBoxListWindow by IceLolly
'Date: 2018.8.10

'Please note that this window is used by Dark¡áComboBox and is NOT suggested to use directly.

Public BoundCtl     As DarkComboBox
Public MaxHeight    As Integer
Public MaxWidth     As Integer
Public PrevIndex    As Integer

Public Sub AddItem(ItemText As String)
    Dim NewIndex    As Integer
    
    NewIndex = Me.labItem.UBound + 1
    Load Me.labItem(NewIndex)                   'Base = 1
    With Me.labItem(NewIndex)
        If NewIndex = 1 Then
            .Top = 0
        Else
            .Top = Me.labItem(NewIndex - 1).Top + Me.labItem(NewIndex - 1).Height
        End If
        .Caption = ItemText
        If .Width > MaxWidth Then
            Dim i As Integer
            
            MaxWidth = .Width + 60
            For i = 0 To NewIndex - 1
                Me.labItem(i).Width = MaxWidth
            Next i
        Else
            .Width = MaxWidth
        End If
        .Visible = True
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        BoundCtl.HideList
    End If
End Sub

Private Sub Form_Load()
    'PrevListWindowProc = SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf ListWindowProc)    '[ToDo]
    PrevIndex = -1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'SetWindowLongA Me.hWnd, GWL_WNDPROC, PrevListWindowProc            '[ToDo]
End Sub

Private Sub labItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i               As Integer
    
    If Index = PrevIndex Then
        Exit Sub
    End If
    For i = 0 To Me.labItem.UBound
        Me.labItem(i).BackColor = RGB(27, 27, 28)
    Next i
    Me.labItem(Index).BackColor = RGB(53, 53, 53)
    PrevIndex = Index
End Sub

Private Sub labItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        BoundCtl.ListIndex = Index
        BoundCtl.HideList
    End If
End Sub

Private Sub tmrCheckFocus_Timer()
    Dim pt      As POINT
    Dim Target  As Long
    
    GetCursorPos pt
    Target = WindowFromPoint(pt.X, pt.Y)
    If Target <> Me.hWnd And Target <> Me.picContainer.hWnd Then
        Dim i As Integer
    
        For i = 0 To Me.labItem.UBound
            Me.labItem(i).BackColor = RGB(27, 27, 28)
        Next i
        PrevIndex = -1
    Else
        ReleaseCapture
    End If
    If GetForegroundWindow() <> Me.hWnd Then
        BoundCtl.HideList
    End If
End Sub

Private Sub VScrollBar_ValueChanged(NewValue As Long)
    Me.picContainer.Top = -NewValue
End Sub

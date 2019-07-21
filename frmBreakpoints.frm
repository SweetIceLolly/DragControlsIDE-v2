VERSION 5.00
Begin VB.Form frmBreakpoints 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "∂œµ„¡–±Ì"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   Icon            =   "frmBreakpoints.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DragControlsIDE.DarkListView lvBreakpoints 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5318
   End
End
Attribute VB_Name = "frmBreakpoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = Lang_Breakpoints_Caption
    Me.lvBreakpoints.AddColumnHeader "ddd"
    Me.lvBreakpoints.AddColumnHeader "yyy"
    Me.lvBreakpoints.AddItem "sdfsaf"
    Me.lvBreakpoints.AddItem "kfckfc"
    Me.lvBreakpoints.AddItem "sdfsaf"
    Me.lvBreakpoints.AddItem "kfckfc"
    Me.lvBreakpoints.AddItem "sdfsaf"
    Me.lvBreakpoints.AddItem "kfckfc"
    Me.lvBreakpoints.SetItemText "sdfsdfsaffsaf", 2, 1
End Sub

Private Sub Form_Resize()
    Me.lvBreakpoints.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

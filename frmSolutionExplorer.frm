VERSION 5.00
Begin VB.Form frmSolutionExplorer 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "工程资源管理器"
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DragControlsIDE.DarkTreeView SolutionTreeView 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5106
   End
End
Attribute VB_Name = "frmSolutionExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'描述:      工程资源管理器，负责显示工程所包含的目录和文件
'作者:      冰棍
'文件:      frmSolutionExplorer.frm
'====================================================

Option Explicit

Public Sub SolutionTreeView_BeginLabelEdit(ByVal hTreeItem As Long, bCancel As Boolean)
    
End Sub

Public Sub SolutionTreeView_Click(bCancel As Boolean)
    
End Sub

Public Sub SolutionTreeView_DoubleClick(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Public Sub SolutionTreeView_EndLabelEdit(ByVal hTreeItem As Long, NewText As String, bCancel As Boolean)
    '如果NewText为vbNullChar，则说明编辑被取消了
End Sub

Public Sub SolutionTreeView_ItemExpanding(ByVal hTreeItem As Long, bCancel As Boolean)

End Sub

Public Sub SolutionTreeView_KeyDown(ByVal KeyCode As Long, ByVal IsLongPress As Boolean)

End Sub

Public Sub SolutionTreeView_KeyUp(ByVal KeyCode As Long)

End Sub

Public Sub SolutionTreeView_MouseDown(ByVal Button As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Public Sub SolutionTreeView_MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Public Sub SolutionTreeView_MouseUp(ByVal Button As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Public Sub SolutionTreeView_RightClick(bCancel As Boolean)

End Sub

Public Sub SolutionTreeView_SelChanged(ByVal hOldTreeItem As Long, ByVal hNewTreeItem As Long)
    
End Sub

Public Sub SolutionTreeView_SelChanging(ByVal hOldTreeItem As Long, ByVal hNewTreeItem As Long, bCancel As Boolean)
    
End Sub

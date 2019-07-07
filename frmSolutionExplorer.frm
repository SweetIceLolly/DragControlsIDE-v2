VERSION 5.00
Begin VB.Form frmSolutionExplorer 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "������Դ������"
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
'����:      ������Դ��������������ʾ������������Ŀ¼���ļ�
'����:      ����
'�ļ�:      frmSolutionExplorer.frm
'====================================================

Option Explicit

Private Sub Form_Resize()
    On Error Resume Next
    Me.SolutionTreeView.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Sub SolutionTreeView_BeginLabelEdit(ByVal hTreeItem As Long, bCancel As Boolean)
    
End Sub

Public Sub SolutionTreeView_Click(bCancel As Boolean)
    
End Sub

Public Sub SolutionTreeView_DoubleClick(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    On Error Resume Next
    Dim CurrSelItem     As Long
    Dim i               As Long
    
    CurrSelItem = Me.SolutionTreeView.GetSelectedItem()                                         '��ȡѡ�������ͼ�б���
    For i = 0 To UBound(TvItemBinding)                                                          '�����б����Ӧ���ļ����
        If CurrSelItem = TvItemBinding(i).TVITEM Then
            If CurrentProject.Files(TvItemBinding(i).FileIndex).TargetWindow Is Nothing Then
                Dim NewCodeWindow   As frmCodeWindow                                                        '�½��Ĵ������
                Dim FileTitle       As String                                                               '�ļ���
                
                Set NewCodeWindow = CreateNewCodeWindow(TvItemBinding(i).FileIndex)                         '�����µĴ��봰�岢���ð󶨵��ļ����
                FileTitle = CurrentProject.Files(TvItemBinding(i).FileIndex).FilePath
                FileTitle = Right(FileTitle, Len(FileTitle) - InStrRev(FileTitle, "\"))                     '��ȡ���ļ���
                NewCodeWindow.Caption = FileTitle
                frmMain.TabBar.AddForm NewCodeWindow
            Else
                CurrentProject.Files(TvItemBinding(i).FileIndex).TargetWindow.Show                          '��ʾ��Ӧ�Ĵ���
                CurrentProject.Files(TvItemBinding(i).FileIndex).TargetWindow.SyntaxEdit.SetFocus
            End If
            Exit For
        End If
    Next i
End Sub

Public Sub SolutionTreeView_EndLabelEdit(ByVal hTreeItem As Long, NewText As String, bCancel As Boolean)
    '���NewTextΪvbNullChar����˵���༭��ȡ����
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

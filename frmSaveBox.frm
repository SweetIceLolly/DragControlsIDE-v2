VERSION 5.00
Begin VB.Form frmSaveBox 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "����"
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H00333333&
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   960
      Width           =   2775
   End
   Begin DragControlsIDE.DarkButton cmdCancel 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ȡ��"
      HasBorder       =   0   'False
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   4320
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      Sizable         =   0   'False
   End
   Begin DragControlsIDE.DarkButton cmdNo 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "��"
      HasBorder       =   0   'False
   End
   Begin DragControlsIDE.DarkButton cmdYes 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "��"
      HasBorder       =   0   'False
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar_NoDrop 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "����"
      MaxButtonEnabled=   0   'False
      MinButtonEnabled=   0   'False
      MaxButtonVisible=   0   'False
      MinButtonVisible=   0   'False
      BindCaption     =   -1  'True
      Picture         =   "frmSaveBox.frx":0000
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ƿ񱣴�������ѡ����ļ���"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   2340
   End
End
Attribute VB_Name = "frmSaveBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'����:      ���洰�ڣ������г����н�Ҫ������ļ����û�����ѡ�񲻱���һЩ�ļ�
'����:      ����
'�ļ�:      frmSaveBox.frm
'====================================================

Option Explicit

'lstFiles.ListIndex����Ӧ��CurrentProject.Files�������������һ��Ԫ�أ�
'�����Ӧ��������-1�����Ӧ���ǵ�ǰ�����ļ�
Dim FileIndexMap()  As Long

Public bBlock       As Boolean                                                  '��������ִ�б��
Public bSaveFlag    As Integer                                                  '��������ǡ�0=��δָ������; 1=����ɹ�; 2=����ʧ��; 3=ȡ��; 4=������

'����:      ��ʼ��FileIndexMap����
Public Sub InitFileIndexMap()
    ReDim FileIndexMap(0)
End Sub

'����:      ��FileIndexMap���������������
'����:      FileIndex: CurrentProject.Files���������������-1�����Ӧ���ǵ�ǰ�����ļ�
Public Sub AddFileIndexMap(FileIndex As Long)
    Dim NewIndex    As Long
    
    NewIndex = UBound(FileIndexMap)
    If FileIndex = -1 Then                                                      '�����ļ�
        Me.lstFiles.AddItem GetFileName(ProjectFilePath)
    Else                                                                        '�����ļ�
        Me.lstFiles.AddItem GetFileName(CurrentProject.Files(FileIndex).FilePath)
    End If
    ReDim Preserve FileIndexMap(NewIndex + 1)
    FileIndexMap(NewIndex) = FileIndex
End Sub

Private Sub cmdCancel_Click()
    bSaveFlag = 3                                                               '���Ϊȡ��
    Unload Me
End Sub

Private Sub cmdNo_Click()
    bSaveFlag = 4                                                               '���Ϊ������
    Unload Me
End Sub

Private Sub cmdYes_Click()
    On Error Resume Next
    Dim i           As Long
    Dim lstIndex    As Long
    'ToDo: detect files of the same name
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE  'ȡ�������ö�����ֹ�ڵ��Ի���
    bSaveFlag = 1                                                               '�ȱ��Ϊ����ʧ��
    For lstIndex = 0 To Me.lstFiles.ListCount - 1                               '�������й�ѡ�˵��ļ�
        If Me.lstFiles.Selected(lstIndex) Then
            If FileIndexMap(lstIndex) = -1 Then                                         '�����ļ�
                Dim ProjectFile_Save    As ProjectFileStruct_Save                           '�����õĹ�����Ϣ�ṹ

                With ProjectFile_Save                                                       '���ƹ�����Ϣ
                    .ProjectName = CurrentProject.ProjectName
                    .ProjectType = CurrentProject.ProjectType
                    ReDim .Files(UBound(CurrentProject.Files))
                    For i = 0 To UBound(.Files)                                                 '�������д����ļ���Ϣ
                        With .Files(i)
                            .FileName = GetFileName(CurrentProject.Files(i).FilePath)
                            .IsHeaderFile = CurrentProject.Files(i).IsHeaderFile
                            .PrevLine = CurrentProject.Files(i).PrevLine
                            .Breakpoints = CurrentProject.Files(i).Breakpoints
                        End With
                    Next i
                End With
                Open ProjectFilePath For Binary As #1                                       '���湤���ļ�
                If Err.Number <> 0 Then                                                     '�����ļ����ܼ���
                    bSaveFlag = 2                                                           '���Ϊ����ʧ��
                    Close #1
                    If NoSkinMsgBox(Lang_SaveBox_SaveFailure_1 & ProjectFilePath & " :" & _
                       Err.Number & " - " & Err.Description & Lang_SaveBox_SaveFailure_2, vbExclamation Or vbYesNo, Lang_Msgbox_Error) = vbNo Then
                        Unload Me
                        Exit Sub
                    End If
                Else                                                                        '�����ļ����Լ���
                    Put #1, , ProjectFile_Save
                    Close #1
                    If Err.Number = 0 Then                                                      '�����ļ��ɹ�
                        CurrentProject.Changed = False                                              '��ǹ����ļ�Ϊ�ѱ���
                    Else                                                                        '�����ļ�ʧ��
                        bSaveFlag = 2                                                               '���Ϊ����ʧ��
                        If NoSkinMsgBox(Lang_SaveBox_SaveFailure_1 & ProjectFilePath & " :" & _
                           Err.Number & " - " & Err.Description & Lang_SaveBox_SaveFailure_2, vbExclamation Or vbYesNo, Lang_Msgbox_Error) = vbNo Then
                            Unload Me
                            Exit Sub
                        End If
                    End If
                End If
            Else                                                                        '�����ļ�
                Open CurrentProject.Files(FileIndexMap(lstIndex)).FilePath For Output As #1
                If Err.Number <> 0 Then                                                     '�����ļ����ܼ���
                    bSaveFlag = 2                                                           '���Ϊ����ʧ��
                    Close #1
                    If NoSkinMsgBox(Lang_SaveBox_SaveFailure_1 & CurrentProject.Files(FileIndexMap(lstIndex)).FilePath & " :" & _
                       Err.Number & " - " & Err.Description & Lang_SaveBox_SaveFailure_2, vbExclamation Or vbYesNo, Lang_Msgbox_Error) = vbNo Then
                        Unload Me
                        Exit Sub
                    End If
                Else                                                                        '�����ļ����Լ���
                    Print #1, CurrentProject.Files(FileIndexMap(lstIndex)).TargetWindow.SyntaxEdit.Text
                    Close #1
                    If Err.Number = 0 Then                                                      '�����ļ��ɹ�
                        CurrentProject.Files(FileIndexMap(lstIndex)).Changed = False                '����ļ�Ϊ�ѱ���
                    Else
                        bSaveFlag = 2                                                               '���Ϊ����ʧ��
                        If NoSkinMsgBox(Lang_SaveBox_SaveFailure_1 & CurrentProject.Files(FileIndexMap(lstIndex)).FilePath & " :" & _
                           Err.Number & " - " & Err.Description & Lang_SaveBox_SaveFailure_2, vbExclamation Or vbYesNo, Lang_Msgbox_Error) = vbNo Then
                            Unload Me
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Next lstIndex
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_SaveBox_Caption
    Me.cmdCancel.Caption = Lang_SaveBox_Cancel
    Me.cmdNo.Caption = Lang_SaveBox_No
    Me.cmdYes.Caption = Lang_SaveBox_Yes
    Me.labTip.Caption = Lang_SaveBox_Prompt
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Enabled = True
    If bBlock = True Then                                                       '���������������ִ�У�˵������ֵ����֮��ᱻʹ�ã�������ȡ�����رգ���ֹ����ֵ��ʧ
        Cancel = 1
        bBlock = False
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.lstFiles.Move 240, Me.labTip.Top + Me.labTip.Height + 240, Me.ScaleWidth - 480, Me.cmdYes.Top - Me.lstFiles.Top - 240
End Sub

Private Sub lstFiles_Click()
    If Me.lstFiles.SelCount = 0 Then                                            '���û��ѡ���ļ����Ͳ������¡��ǡ�
        Me.cmdYes.Enabled = False
    Else
        Me.cmdYes.Enabled = True
    End If
End Sub

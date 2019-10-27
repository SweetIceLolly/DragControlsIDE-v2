VERSION 5.00
Begin VB.Form frmSaveBox 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "保存"
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
      Caption         =   "取消"
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
      Caption         =   "否"
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
      Caption         =   "是"
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
      Caption         =   "保存"
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
      Caption         =   "是否保存下列所选择的文件？"
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
'描述:      保存窗口，用于列出所有将要保存的文件，用户可以选择不保存一些文件
'作者:      冰棍
'文件:      frmSaveBox.frm
'====================================================

Option Explicit

'lstFiles.ListIndex所对应的CurrentProject.Files索引（弃用最后一个元素）
'如果对应的索引是-1，则对应的是当前工程文件
Dim FileIndexMap()  As Long

Public bBlock       As Boolean                                                  '阻塞代码执行标记
Public bSaveFlag    As Integer                                                  '保存结果标记。0=尚未指定操作; 1=保存成功; 2=保存失败; 3=取消; 4=不保存

'描述:      初始化FileIndexMap数组
Public Sub InitFileIndexMap()
    ReDim FileIndexMap(0)
End Sub

'描述:      往FileIndexMap数组里面添加索引
'参数:      FileIndex: CurrentProject.Files索引。如果索引是-1，则对应的是当前工程文件
Public Sub AddFileIndexMap(FileIndex As Long)
    Dim NewIndex    As Long
    
    NewIndex = UBound(FileIndexMap)
    If FileIndex = -1 Then                                                      '工程文件
        Me.lstFiles.AddItem GetFileName(ProjectFilePath)
    Else                                                                        '代码文件
        Me.lstFiles.AddItem GetFileName(CurrentProject.Files(FileIndex).FilePath)
    End If
    ReDim Preserve FileIndexMap(NewIndex + 1)
    FileIndexMap(NewIndex) = FileIndex
End Sub

Private Sub cmdCancel_Click()
    bSaveFlag = 3                                                               '标记为取消
    Unload Me
End Sub

Private Sub cmdNo_Click()
    bSaveFlag = 4                                                               '标记为不保存
    Unload Me
End Sub

Private Sub cmdYes_Click()
    On Error Resume Next
    Dim i           As Long
    Dim lstIndex    As Long
    
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE  '取消窗口置顶，防止遮挡对话框
    bSaveFlag = 1                                                               '先标记为保存失败
    For lstIndex = 0 To Me.lstFiles.ListCount - 1                               '保存所有勾选了的文件
        If Me.lstFiles.Selected(lstIndex) Then
            If FileIndexMap(lstIndex) = -1 Then                                         '工程文件
                Dim ProjectFile_Save    As ProjectFileStruct_Save                           '保存用的工程信息结构

                With ProjectFile_Save                                                       '复制工程信息
                    .ProjectName = CurrentProject.ProjectName
                    .ProjectType = CurrentProject.ProjectType
                    ReDim .Files(UBound(CurrentProject.Files))
                    For i = 0 To UBound(.Files)                                                 '复制所有代码文件信息
                        With .Files(i)
                            .FileName = GetFileName(CurrentProject.Files(i).FilePath)
                            .PrevLine = CurrentProject.Files(i).PrevLine
                            .Breakpoints = CurrentProject.Files(i).Breakpoints
                            .FolderIndex = CurrentProject.Files(i).FolderIndex
                        End With
                    Next i
                    .Folders = CurrentProject.Folders                                           '复制所有文件夹信息
                End With
                Open ProjectFilePath For Binary As #1                                       '保存工程文件
                If Err.Number <> 0 Then                                                     '保存文件不能继续
                    bSaveFlag = 2                                                           '标记为保存失败
                    Close #1
                    If NoSkinMsgBox(Lang_SaveBox_SaveFailure_1 & ProjectFilePath & " :" & _
                       Err.Number & " - " & Err.Description & Lang_SaveBox_SaveFailure_2, vbExclamation Or vbYesNo, Lang_Msgbox_Error) = vbNo Then
                        Unload Me
                        Exit Sub
                    End If
                Else                                                                        '保存文件可以继续
                    Put #1, , ProjectFile_Save
                    Close #1
                    If Err.Number = 0 Then                                                      '保存文件成功
                        CurrentProject.Changed = False                                              '标记工程文件为已保存
                    Else                                                                        '保存文件失败
                        bSaveFlag = 2                                                               '标记为保存失败
                        If NoSkinMsgBox(Lang_SaveBox_SaveFailure_1 & ProjectFilePath & " :" & _
                           Err.Number & " - " & Err.Description & Lang_SaveBox_SaveFailure_2, vbExclamation Or vbYesNo, Lang_Msgbox_Error) = vbNo Then
                            Unload Me
                            Exit Sub
                        End If
                    End If
                End If
            Else                                                                        '代码文件
                Open CurrentProject.Files(FileIndexMap(lstIndex)).FilePath For Output As #1
                If Err.Number <> 0 Then                                                     '保存文件不能继续
                    bSaveFlag = 2                                                           '标记为保存失败
                    Close #1
                    If NoSkinMsgBox(Lang_SaveBox_SaveFailure_1 & CurrentProject.Files(FileIndexMap(lstIndex)).FilePath & " :" & _
                       Err.Number & " - " & Err.Description & Lang_SaveBox_SaveFailure_2, vbExclamation Or vbYesNo, Lang_Msgbox_Error) = vbNo Then
                        Unload Me
                        Exit Sub
                    End If
                Else                                                                        '保存文件可以继续
                    Print #1, CurrentProject.Files(FileIndexMap(lstIndex)).TargetWindow.SyntaxEdit.Text
                    Close #1
                    If Err.Number = 0 Then                                                      '保存文件成功
                        CurrentProject.Files(FileIndexMap(lstIndex)).Changed = False                '标记文件为已保存
                    Else
                        bSaveFlag = 2                                                               '标记为保存失败
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
    If bBlock = True Then                                                       '如果正在阻塞代码执行，说明返回值可能之后会被使用，所以先取消掉关闭，防止返回值丢失
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
    If Me.lstFiles.SelCount = 0 Then                                            '如果没有选择文件，就不给按下“是”
        Me.cmdYes.Enabled = False
    Else
        Me.cmdYes.Enabled = True
    End If
End Sub

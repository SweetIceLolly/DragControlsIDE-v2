VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmCodeWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "���봰��"
   ClientHeight    =   5175
   ClientLeft      =   3540
   ClientTop       =   3060
   ClientWidth     =   8865
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmCodeWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   8865
   Begin XtremeSyntaxEdit.SyntaxEdit SyntaxEdit 
      Height          =   1935
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
      _Version        =   983043
      _ExtentX        =   5318
      _ExtentY        =   3413
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   0   'False
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin VB.Timer tmrUpdateBreakpoints 
      Interval        =   50
      Left            =   6960
      Top             =   4560
   End
   Begin VB.PictureBox picSelMargin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00333333&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   240
      Picture         =   "frmCodeWindow.frx":1BCC2
      ScaleHeight     =   1935
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   1320
      Width           =   255
   End
   Begin DragControlsIDE.DarkComboBox comObject 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      Items0          =   ""
      ITEM_COUNT      =   0
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   7560
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      Thickness       =   4
      MinWidth        =   150
      MinHeight       =   100
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8865
      _ExtentX        =   15637
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
      Caption         =   "���봰��"
      MaxButtonVisible=   0   'False
      MinButtonVisible=   0   'False
      BindCaption     =   -1  'True
      Picture         =   "frmCodeWindow.frx":1C04C
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorderSizer 
      Left            =   8280
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      Thickness       =   3
      FocusedColor    =   3157293
      NotFocusedColor =   3157293
      MinWidth        =   150
      MinHeight       =   100
   End
   Begin DragControlsIDE.DarkComboBox comEvent 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   660
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      Items0          =   ""
      ITEM_COUNT      =   0
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCodeWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'����:      ���봰�ڣ�����󲿷ֵĴ�����صĹ���
'����:      ����
'�ļ�:      frmCodeWindow.frm
'====================================================

Option Explicit

Public WindowObj    As Object                                                       '��������
Public FileIndex    As Long                                                         '��CurrentProject.Files��Ӧ���ļ����

Dim BreakpointPic   As StdPicture                                                   '�ϵ�ͼƬ
Dim RowHeight       As Integer                                                      '�����еĸ߶ȣ����ڼ���ϵ��ͼλ�ã�

'����:      ����ͨ���������������ÿ�д���ĸ߶�
Public Sub ReCalcRowHeight()
    Set Me.picSelMargin.Font = Me.SyntaxEdit.Font
    RowHeight = Me.picSelMargin.TextHeight("#")
End Sub

'����:      �ػ����еĶϵ�
Public Sub RedrawBreakpoints()
    Dim lnStart     As Long, lnEnd      As Long, ln         As Long                 '���ӵĵ�һ��; ���ӵ����һ��; �ϵ��Ӧ����
    Dim i           As Long
    
    Me.picSelMargin.Cls                                                             '��ջ���
    lnStart = Me.SyntaxEdit.TopRow                                                  '��ȡ��ǰ���ӵĵ�һ��
    lnEnd = lnStart + Me.SyntaxEdit.Height / RowHeight                              '�����ı���ĸ߶Ⱥ�ÿ�еĸ߶�����ı�����װ�¶����У��Ӷ��õ����ӵ����һ��
    If lnEnd > Me.SyntaxEdit.RowsCount Then                                         '������ӵ����һ�г������ı������������ȡ������
        lnEnd = Me.SyntaxEdit.RowsCount
    End If
    For i = 0 To UBound(CurrentProject.Files(FileIndex).Breakpoints)                '������ǰ�ļ��Ķϵ㣬������ڿ��ӵ�������Χ�ڵľͻ�����
        ln = CurrentProject.Files(FileIndex).Breakpoints(i).CodeLn
        If ln >= lnStart And ln <= lnEnd Then
            Me.picSelMargin.PaintPicture BreakpointPic, 0, RowHeight * (ln - lnStart), 240, 240
        End If
    Next i
End Sub

Private Sub DarkTitleBar_GotFocus()
    On Error Resume Next
    
    Me.SyntaxEdit.SetFocus
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_CodeWindow_Caption
    Me.DarkTitleBar.MaxButtonVisible = True
    Me.DarkTitleBar.MinButtonVisible = True
    
    Set BreakpointPic = Me.picSelMargin.Picture                                                                         '���öϵ�ͼƬ
    Set Me.picSelMargin.Picture = Nothing
    Call ReCalcRowHeight                                                                                                '���¼�������и߶�
    
    '���ô��������
    Me.DarkTitleBar.Top = Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelY
    Me.picSelMargin.Move Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelX, Me.SyntaxEdit.Top, 300, Me.SyntaxEdit.Height
    Me.SyntaxEdit.Move Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelX + Me.picSelMargin.Width, _
        Me.DarkTitleBar.Height + Me.comObject.Height + 240 + Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelY
    Me.SyntaxEdit.PaintManager.BackColor = RGB(28, 28, 28)
    Me.SyntaxEdit.PaintManager.LineNumberBackColor = RGB(28, 28, 28)
    Me.SyntaxEdit.PaintManager.LineNumberTextColor = RGB(86, 156, 214)
    Me.SyntaxEdit.DataManager.FileExt = ".cpp"
    Me.SyntaxEdit.ConfigFile = App.Path & "\SyntaxEdit.ini"
    
    '���ô������໯������������⼰�����������Ҽ��ر�
    Dim lpObj               As Long                                                                                     'ָ�򴰿���������ָ��
    Set WindowObj = Me
    lpObj = ObjPtr(WindowObj)                                                                                           '��ȡָ�򴰿���������ָ��
    SetPropA Me.hWnd, "WindowObj", lpObj                                                                                '��¼���ڵ������ַ�������໯ж�ش�����
    'SetPropA Me.hWnd, "PrevWndProc", SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf MainWindowMaximizeCloseFixProc)    '[ToDo]

    '���ô��������໯��ʹ���ػ��ʱ���ܹ��ػ�ϵ�
    Dim RealSyntaxEdit      As Long                                                                                     '�������ʵ��hWnd
    
    RealSyntaxEdit = FindWindowExA(Me.SyntaxEdit.hWnd, 0, "CodejockSyntaxEditor", vbNullString)                         '�������ʵֻ��һ���ǣ�������Ǹ����ڲ��������Ĵ���򴰿�
    SetPropA RealSyntaxEdit, "FileIndex", FileIndex
    SetPropA RealSyntaxEdit, "PrevWndProc", SetWindowLongA(RealSyntaxEdit, GWL_WNDPROC, AddressOf EditBreakpointsRedrawProc)    '[ToDo]
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsExiting Then
        '�ָ��������໯
        SetWindowLongA Me.hWnd, GWL_WNDPROC, GetPropA(Me.hWnd, "PrevWndProc")
        SetWindowLongA Me.SyntaxEdit.hWnd, GWL_WNDPROC, GetPropA(Me.SyntaxEdit.hWnd, "PrevWndProc")
    Else
        Cancel = 1
        Me.Hide
    End If
End Sub

Public Sub Form_Resize()
    On Error Resume Next
    
    '���ݱ������Ƿ���ʾ�������ؼ�λ��
    If Me.DarkTitleBar.Visible = True Then
        Me.comObject.Top = Me.DarkTitleBar.Height + 165
        Me.comEvent.Top = Me.comObject.Top
        Me.SyntaxEdit.Top = Me.comEvent.Top + Me.comEvent.Height + 240
    Else
        Me.comObject.Top = 120
        Me.comEvent.Top = 120
        Me.SyntaxEdit.Top = 120 + Me.comObject.Height + 120
    End If
    Me.picSelMargin.Top = Me.SyntaxEdit.Top
    
    '���ô�����С
    Me.SyntaxEdit.Width = Me.ScaleWidth - Me.SyntaxEdit.Left - Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelX
    Me.SyntaxEdit.Height = Me.ScaleHeight - Me.SyntaxEdit.Top - Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelY
    Me.picSelMargin.Height = Me.SyntaxEdit.Height
    
    '������Ͽ��С��λ��
    Me.comObject.Width = (Me.ScaleWidth - 480) / 2
    Me.comEvent.Width = Me.comObject.Width
    Me.comEvent.Left = Me.comObject.Width + 360
End Sub

Private Sub picSelMargin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim CurrRow         As Long, CurrCol        As Long                                 '��������Ӧ�Ĵ����С��С���������û�õģ�������������ؼ����Ҫ�Ҵ��������...
    Dim BreakpointCount As Long                                                         'UBound(.Breakpoints)��ʵ�ʶϵ����� - 1
    Dim i               As Long
    
    Me.SyntaxEdit.RowColCodeFromPoint X, Y / Screen.TwipsPerPixelY, CurrRow, CurrCol    '��ȡ��������Ӧ����
    Me.SyntaxEdit.SetFocus
    
    With CurrentProject.Files(FileIndex)
        BreakpointCount = UBound(.Breakpoints)
        For i = 0 To BreakpointCount                                                    '���Ҷ�Ӧ�Ķϵ�
            If .Breakpoints(i).CodeLn = CurrRow Then                                        '������ҵ���Ӧ�Ķϵ��ɾ��
                Dim j               As Long
                
                frmBreakpoints.lvBreakpoints.DeleteItem .Breakpoints(i).ListViewIndex           '��ListView�Ƴ���Ӧ���б���
                For j = 0 To BreakpointCount                                                    '�������и��б�������Ӧ�Ķϵ㣬������������Ӧ���б������ - 1
                    If .Breakpoints(j).ListViewIndex > .Breakpoints(i).ListViewIndex Then
                        .Breakpoints(j).ListViewIndex = .Breakpoints(j).ListViewIndex - 1
                    End If
                Next j
                
                If i < BreakpointCount Then                                                     '������滹�б�Ķϵ���Ϣ�Ͱ�������ǰ��
                    CopyMemory .Breakpoints(i), .Breakpoints(i + 1), LenB(.Breakpoints(0)) * (BreakpointCount - i)
                End If
                ReDim Preserve .Breakpoints(BreakpointCount - 1)                                '��С�ϵ�����
                Call RedrawBreakpoints                                                          '�ػ����жϵ�
                Exit Sub
            End If
        Next i
        
        '��������ҵ���Ӧ�Ķϵ�����
        ReDim Preserve .Breakpoints(BreakpointCount + 1)                                '����ϵ�����
        .Breakpoints(BreakpointCount).CodeLn = CurrRow                                  '���öϵ��Ӧ�������ͼ���״̬
        .Breakpoints(BreakpointCount).Enabled = True
        .Breakpoints(BreakpointCount).ListViewIndex = frmBreakpoints.lvBreakpoints.AddItem(GetFileName(.FilePath))
        frmBreakpoints.lvBreakpoints.SetItemText CStr(CurrRow), .Breakpoints(BreakpointCount).ListViewIndex, 1
        frmBreakpoints.lvBreakpoints.SetItemChecked .Breakpoints(BreakpointCount).ListViewIndex, True
        Call RedrawBreakpoints                                                          '�ػ����жϵ�
    End With
End Sub

Private Sub SyntaxEdit_TextChanged(ByVal nRowFrom As Long, ByVal nRowTo As Long, ByVal nActions As Long)
    CurrentProject.Files(FileIndex).Changed = True                                  '����������һ�����ģ��Ͱ��ļ���Ϊ������
    '------------------------------------------------------
    
    Dim nLinesChanged   As Long                                                     '�仯������
    Dim SelStartRow     As Long                                                     'ѡ����ı�����ʼ��
    Dim SelEndRow       As Long                                                     'ѡ����ı��Ľ�����
    Dim i               As Long
    Dim j               As Long
    
    If nRowTo <> nRowFrom Then                                                      '������������˱仯
        nLinesChanged = nRowTo - nRowFrom                                               '���������ı仯
        Select Case nActions                                                            '��һЩ�������д���
            Case 6                                                                          'ɾ���������˸����ɾ���������еȣ�
                nLinesChanged = -nLinesChanged
            
            Case 775, 518, 261                                                              '�������ظ�
                nLinesChanged = 0
                
        End Select
    End If
    If nLinesChanged = 0 Then                                                       '������������˱仯�ż��ϵ���û���ܵ�Ӱ��
        Exit Sub
    End If
    SelStartRow = Me.SyntaxEdit.Selection.Start.Row
    SelEndRow = Me.SyntaxEdit.Selection.End.Row
    
    With CurrentProject.Files(FileIndex)
        For i = UBound(.Breakpoints) - 1 To 0 Step -1                                       '�����ϵ��б�ɾ���漰�Ķϵ㣬�����������ϵ��λ��
            If nLinesChanged < 0 And _
               ((SelEndRow <= .Breakpoints(i).CodeLn And .Breakpoints(i).CodeLn <= SelStartRow And SelEndRow < SelStartRow) Or _
               (SelStartRow <= .Breakpoints(i).CodeLn And .Breakpoints(i).CodeLn <= SelEndRow And SelStartRow <= SelEndRow)) Then
                '�ϵ�λ�ڱ�ɾ�������м䣨SelEndRow �� SelStartRow ���Ի���λ�ã���Ϊ�û����ĵķ�����Բ�һ����
                ' ...
                ' SelEndRow   -----  ��
                ' ...                ��
                '  .CodeLn    -----  �� ���м�Ķϵ㽫��ɾ��
                ' ...                ��
                ' SelStartRow -----  ��
                ' ...
                '=====================
                'ɾ���ϵ㡣����Ĵ���������picSelMargin_MouseDown��ɾ���ϵ�Ĵ���
                frmBreakpoints.lvBreakpoints.DeleteItem .Breakpoints(i).ListViewIndex       '��ListView�Ƴ���Ӧ���б���
                For j = 0 To UBound(.Breakpoints)                                           '�������и��б�������Ӧ�Ķϵ㣬������������Ӧ���б������ - 1
                    If .Breakpoints(j).ListViewIndex > .Breakpoints(i).ListViewIndex Then
                        .Breakpoints(j).ListViewIndex = .Breakpoints(j).ListViewIndex - 1
                    End If
                Next j
                
                If i < UBound(.Breakpoints) Then                                            '������滹�б�Ķϵ���Ϣ�Ͱ�������ǰ��
                    CopyMemory .Breakpoints(i), .Breakpoints(i + 1), LenB(.Breakpoints(0)) * (UBound(.Breakpoints) - i)
                End If
                ReDim Preserve .Breakpoints(UBound(.Breakpoints) - 1)                       '��С�ϵ�����
            ElseIf .Breakpoints(i).CodeLn > nRowFrom Then
                '�ϵ�λ�ڷ������ĵ��к���
                ' ...
                ' nRowFrom -----
                ' ...               ��
                ' .CodeLn -----     �� ��nRowFrom����Ķϵ�����Ӧ���кŽ����޸�
                ' ...               ��
                '=====================
                .Breakpoints(i).CodeLn = .Breakpoints(i).CodeLn + nLinesChanged
                frmBreakpoints.lvBreakpoints.SetItemText CStr(.Breakpoints(i).CodeLn), .Breakpoints(i).ListViewIndex, 1
            End If
        Next i
    End With
    
    Call RedrawBreakpoints                                                          '�ػ����жϵ�
    bpRedrawFileIndex = -1                                                          '�ü�ʱ�����ػ���
End Sub

Private Sub tmrUpdateBreakpoints_Timer()
    If bpRedrawFileIndex = FileIndex Then
        Call RedrawBreakpoints
        bpRedrawFileIndex = -1
    End If
End Sub

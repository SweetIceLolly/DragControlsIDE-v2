VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "CO7FCA~1.OCX"
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "COCEAE~1.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "COE2B7~1.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "Microsoft YaHei UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   13140
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00302D2D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   13140
      TabIndex        =   2
      Top             =   810
      Width           =   13140
      Begin XtremeSyntaxEdit.SyntaxEdit SyntaxEdit 
         Height          =   855
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1455
         _Version        =   983043
         _ExtentX        =   2566
         _ExtentY        =   1508
         _StockProps     =   84
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft YaHei UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableSyntaxColorization=   -1  'True
         ShowLineNumbers =   -1  'True
         ShowSelectionMargin=   -1  'True
         ShowScrollBarVert=   -1  'True
         ShowScrollBarHorz=   -1  'True
         EnableVirtualSpace=   0   'False
         EnableAutoIndent=   -1  'True
         ShowWhiteSpace  =   0   'False
         ShowCollapsibleNodes=   -1  'True
         AutoCompleteWndWidth=   160
         EnableEditAccelerators=   -1  'True
      End
      Begin VB.PictureBox picScrollBarCover 
         Appearance      =   0  'Flat
         BackColor       =   &H00302D2D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9720
         ScaleHeight     =   255
         ScaleMode       =   0  'User
         ScaleWidth      =   255
         TabIndex        =   4
         Top             =   1800
         Width           =   255
      End
   End
   Begin �Ͽؼ���UI.DarkMenu DarkMenu 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   495
      Width           =   13140
      _ExtentX        =   23178
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MENU_ITEM_COUNT =   54
      LEVELS_COUNT    =   54
      LEVELS_2        =   1
      LEVELS_3        =   1
      LEVELS_4        =   1
      LEVELS_5        =   1
      LEVELS_6        =   1
      LEVELS_7        =   1
      LEVELS_9        =   1
      LEVELS_10       =   1
      LEVELS_11       =   1
      LEVELS_12       =   1
      LEVELS_13       =   1
      LEVELS_14       =   1
      LEVELS_15       =   1
      LEVELS_16       =   1
      LEVELS_17       =   1
      LEVELS_18       =   1
      LEVELS_19       =   1
      LEVELS_20       =   1
      LEVELS_21       =   1
      LEVELS_22       =   1
      LEVELS_23       =   1
      LEVELS_24       =   1
      LEVELS_25       =   1
      LEVELS_26       =   1
      LEVELS_27       =   1
      LEVELS_29       =   1
      LEVELS_30       =   1
      LEVELS_31       =   1
      LEVELS_32       =   1
      LEVELS_33       =   1
      LEVELS_34       =   1
      LEVELS_36       =   1
      LEVELS_37       =   1
      LEVELS_39       =   1
      LEVELS_40       =   1
      LEVELS_41       =   1
      LEVELS_42       =   1
      LEVELS_43       =   1
      LEVELS_44       =   1
      LEVELS_46       =   1
      LEVELS_47       =   1
      LEVELS_48       =   1
      LEVELS_49       =   1
      LEVELS_50       =   1
      LEVELS_51       =   1
      LEVELS_53       =   1
      LEVELS_54       =   1
      MenuID_1        =   0
      MenuText_1      =   "�ļ�"
      MenuVisible_1   =   -1  'True
      SUBMENU_ITEM_COUNT_1=   6
      SubMenuID_1_0   =   0
      SubMenuText_1_1 =   "�½�                   "
      SubMenuID_1_1   =   2
      SubMenuText_1_2 =   "����"
      SubMenuID_1_2   =   3
      SubMenuText_1_3 =   "����"
      SubMenuID_1_3   =   4
      SubMenuText_1_4 =   "���Ϊ"
      SubMenuID_1_4   =   5
      SubMenuText_1_5 =   "-"
      SubMenuID_1_5   =   6
      SubMenuText_1_6 =   "�˳�"
      SubMenuID_1_6   =   7
      MenuID_2        =   1
      MenuText_2      =   "�½�                   "
      MenuVisible_2   =   -1  'True
      SubMenuID_2_0   =   0
      MenuID_3        =   2
      MenuText_3      =   "����"
      MenuVisible_3   =   -1  'True
      SubMenuID_3_0   =   0
      MenuID_4        =   3
      MenuText_4      =   "����"
      MenuVisible_4   =   -1  'True
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "���Ϊ"
      MenuVisible_5   =   -1  'True
      SubMenuID_5_0   =   0
      MenuID_6        =   5
      MenuText_6      =   "-"
      MenuVisible_6   =   -1  'True
      SubMenuID_6_0   =   0
      MenuID_7        =   6
      MenuText_7      =   "�˳�"
      MenuVisible_7   =   -1  'True
      SubMenuID_7_0   =   0
      MenuID_8        =   7
      MenuText_8      =   "�༭"
      MenuVisible_8   =   -1  'True
      SUBMENU_ITEM_COUNT_8=   19
      SubMenuID_8_0   =   0
      SubMenuText_8_1 =   "����                        "
      SubMenuID_8_1   =   9
      SubMenuText_8_2 =   "�ظ�"
      SubMenuID_8_2   =   10
      SubMenuText_8_3 =   "-"
      SubMenuID_8_3   =   11
      SubMenuText_8_4 =   "����"
      SubMenuID_8_4   =   12
      SubMenuText_8_5 =   "����"
      SubMenuID_8_5   =   13
      SubMenuText_8_6 =   "ճ��"
      SubMenuID_8_6   =   14
      SubMenuText_8_7 =   "ȫѡ"
      SubMenuID_8_7   =   15
      SubMenuText_8_8 =   "ɾ����"
      SubMenuID_8_8   =   16
      SubMenuText_8_9 =   "-"
      SubMenuID_8_9   =   17
      SubMenuText_8_10=   "����"
      SubMenuID_8_10  =   18
      SubMenuText_8_11=   "�滻"
      SubMenuID_8_11  =   19
      SubMenuText_8_12=   "-"
      SubMenuID_8_12  =   20
      SubMenuText_8_13=   "��������"
      SubMenuID_8_13  =   21
      SubMenuText_8_14=   "��������"
      SubMenuID_8_14  =   22
      SubMenuText_8_15=   "-"
      SubMenuID_8_15  =   23
      SubMenuText_8_16=   "���/�Ƴ��ϵ�"
      SubMenuID_8_16  =   24
      SubMenuText_8_17=   "������жϵ�"
      SubMenuID_8_17  =   25
      SubMenuText_8_18=   "-"
      SubMenuID_8_18  =   26
      SubMenuText_8_19=   "��ת����"
      SubMenuID_8_19  =   27
      MenuID_9        =   8
      MenuText_9      =   "����                        "
      MenuVisible_9   =   -1  'True
      SubMenuID_9_0   =   0
      MenuID_10       =   9
      MenuText_10     =   "�ظ�"
      MenuVisible_10  =   -1  'True
      SubMenuID_10_0  =   0
      MenuID_11       =   10
      MenuText_11     =   "-"
      MenuVisible_11  =   -1  'True
      SubMenuID_11_0  =   0
      MenuID_12       =   11
      MenuText_12     =   "����"
      MenuVisible_12  =   -1  'True
      SubMenuID_12_0  =   0
      MenuID_13       =   12
      MenuText_13     =   "����"
      MenuVisible_13  =   -1  'True
      SubMenuID_13_0  =   0
      MenuID_14       =   13
      MenuText_14     =   "ճ��"
      MenuVisible_14  =   -1  'True
      SubMenuID_14_0  =   0
      MenuID_15       =   14
      MenuText_15     =   "ȫѡ"
      MenuVisible_15  =   -1  'True
      SubMenuID_15_0  =   0
      MenuID_16       =   15
      MenuText_16     =   "ɾ����"
      MenuVisible_16  =   -1  'True
      SubMenuID_16_0  =   0
      MenuID_17       =   16
      MenuText_17     =   "-"
      MenuVisible_17  =   -1  'True
      SubMenuID_17_0  =   0
      MenuID_18       =   17
      MenuText_18     =   "����"
      MenuVisible_18  =   -1  'True
      SubMenuID_18_0  =   0
      MenuID_19       =   18
      MenuText_19     =   "�滻"
      MenuVisible_19  =   -1  'True
      SubMenuID_19_0  =   0
      MenuID_20       =   19
      MenuText_20     =   "-"
      MenuVisible_20  =   -1  'True
      SubMenuID_20_0  =   0
      MenuID_21       =   20
      MenuText_21     =   "��������"
      MenuVisible_21  =   -1  'True
      SubMenuID_21_0  =   0
      MenuID_22       =   21
      MenuText_22     =   "��������"
      MenuVisible_22  =   -1  'True
      SubMenuID_22_0  =   0
      MenuID_23       =   22
      MenuText_23     =   "-"
      MenuVisible_23  =   -1  'True
      SubMenuID_23_0  =   0
      MenuID_24       =   23
      MenuText_24     =   "���/�Ƴ��ϵ�"
      MenuVisible_24  =   -1  'True
      SubMenuID_24_0  =   0
      MenuID_25       =   24
      MenuText_25     =   "������жϵ�"
      MenuVisible_25  =   -1  'True
      SubMenuID_25_0  =   0
      MenuID_26       =   25
      MenuText_26     =   "-"
      MenuVisible_26  =   -1  'True
      SubMenuID_26_0  =   0
      MenuID_27       =   26
      MenuText_27     =   "��ת����"
      MenuVisible_27  =   -1  'True
      SubMenuID_27_0  =   0
      MenuID_28       =   27
      MenuText_28     =   "��ͼ"
      MenuVisible_28  =   -1  'True
      SUBMENU_ITEM_COUNT_28=   6
      SubMenuID_28_0  =   0
      SubMenuText_28_1=   "������                    "
      SubMenuID_28_1  =   29
      SubMenuText_28_2=   "�ؼ���"
      SubMenuID_28_2  =   30
      SubMenuText_28_3=   "���Ա�"
      SubMenuID_28_3  =   31
      SubMenuText_28_4=   "������"
      SubMenuID_28_4  =   32
      SubMenuText_28_5=   "�ϵ��б����"
      SubMenuID_28_5  =   33
      SubMenuText_28_6=   "�����б����"
      SubMenuID_28_6  =   34
      MenuID_29       =   28
      MenuText_29     =   "������                    "
      MenuVisible_29  =   -1  'True
      SubMenuID_29_0  =   0
      MenuID_30       =   29
      MenuText_30     =   "�ؼ���"
      MenuVisible_30  =   -1  'True
      SubMenuID_30_0  =   0
      MenuID_31       =   30
      MenuText_31     =   "���Ա�"
      MenuVisible_31  =   -1  'True
      SubMenuID_31_0  =   0
      MenuID_32       =   31
      MenuText_32     =   "������"
      MenuVisible_32  =   -1  'True
      SubMenuID_32_0  =   0
      MenuID_33       =   32
      MenuText_33     =   "�ϵ��б����"
      MenuVisible_33  =   -1  'True
      SubMenuID_33_0  =   0
      MenuID_34       =   33
      MenuText_34     =   "�����б����"
      MenuVisible_34  =   -1  'True
      SubMenuID_34_0  =   0
      MenuID_35       =   34
      MenuText_35     =   "����"
      MenuVisible_35  =   -1  'True
      SUBMENU_ITEM_COUNT_35=   2
      SubMenuID_35_0  =   0
      SubMenuText_35_1=   "���ɴ����ļ�            "
      SubMenuID_35_1  =   36
      SubMenuText_35_2=   "���ɿ�ִ���ļ�"
      SubMenuID_35_2  =   37
      MenuID_36       =   35
      MenuText_36     =   "���ɴ����ļ�            "
      MenuVisible_36  =   -1  'True
      SubMenuID_36_0  =   0
      MenuID_37       =   36
      MenuText_37     =   "���ɿ�ִ���ļ�"
      MenuVisible_37  =   -1  'True
      SubMenuID_37_0  =   0
      MenuID_38       =   37
      MenuText_38     =   "����"
      MenuVisible_38  =   -1  'True
      SUBMENU_ITEM_COUNT_38=   6
      SubMenuID_38_0  =   0
      SubMenuText_38_1=   "����                     "
      SubMenuID_38_1  =   39
      SubMenuText_38_2=   "�ж�"
      SubMenuID_38_2  =   40
      SubMenuText_38_3=   "ֹͣ"
      SubMenuID_38_3  =   41
      SubMenuText_38_4=   "-"
      SubMenuID_38_4  =   42
      SubMenuText_38_5=   "�����ִ��"
      SubMenuID_38_5  =   43
      SubMenuText_38_6=   "�����ִ��"
      SubMenuID_38_6  =   44
      MenuID_39       =   38
      MenuText_39     =   "����                     "
      MenuVisible_39  =   -1  'True
      SubMenuID_39_0  =   0
      MenuID_40       =   39
      MenuText_40     =   "�ж�"
      MenuVisible_40  =   -1  'True
      SubMenuID_40_0  =   0
      MenuID_41       =   40
      MenuText_41     =   "ֹͣ"
      MenuVisible_41  =   -1  'True
      SubMenuID_41_0  =   0
      MenuID_42       =   41
      MenuText_42     =   "-"
      MenuVisible_42  =   -1  'True
      SubMenuID_42_0  =   0
      MenuID_43       =   42
      MenuText_43     =   "�����ִ��"
      MenuVisible_43  =   -1  'True
      SubMenuID_43_0  =   0
      MenuID_44       =   43
      MenuText_44     =   "�����ִ��"
      MenuVisible_44  =   -1  'True
      SubMenuID_44_0  =   0
      MenuID_45       =   44
      MenuText_45     =   "����"
      MenuVisible_45  =   -1  'True
      SUBMENU_ITEM_COUNT_45=   6
      SubMenuID_45_0  =   0
      SubMenuText_45_1=   "���ڹ���                    "
      SubMenuID_45_1  =   46
      SubMenuText_45_2=   "��Ϣ����"
      SubMenuID_45_2  =   47
      SubMenuText_45_3=   "����"
      SubMenuID_45_3  =   48
      SubMenuText_45_4=   "û��ñ�Ĺ���"
      SubMenuID_45_4  =   49
      SubMenuText_45_5=   "-"
      SubMenuID_45_5  =   50
      SubMenuText_45_6=   "����"
      SubMenuID_45_6  =   51
      MenuID_46       =   45
      MenuText_46     =   "���ڹ���                    "
      MenuVisible_46  =   -1  'True
      SubMenuID_46_0  =   0
      MenuID_47       =   46
      MenuText_47     =   "��Ϣ����"
      MenuVisible_47  =   -1  'True
      SubMenuID_47_0  =   0
      MenuID_48       =   47
      MenuText_48     =   "����"
      MenuVisible_48  =   -1  'True
      SubMenuID_48_0  =   0
      MenuID_49       =   48
      MenuText_49     =   "û��ñ�Ĺ���"
      MenuVisible_49  =   -1  'True
      SubMenuID_49_0  =   0
      MenuID_50       =   49
      MenuText_50     =   "-"
      MenuVisible_50  =   -1  'True
      SubMenuID_50_0  =   0
      MenuID_51       =   50
      MenuText_51     =   "����"
      MenuVisible_51  =   -1  'True
      SubMenuID_51_0  =   0
      MenuID_52       =   51
      MenuText_52     =   "����"
      MenuVisible_52  =   -1  'True
      SUBMENU_ITEM_COUNT_52=   2
      SubMenuID_52_0  =   0
      SubMenuText_52_1=   "�����ĵ�                     "
      SubMenuID_52_1  =   53
      SubMenuText_52_2=   "�����Ͽؼ���"
      SubMenuID_52_2  =   54
      MenuID_53       =   52
      MenuText_53     =   "�����ĵ�                     "
      MenuVisible_53  =   -1  'True
      SubMenuID_53_0  =   0
      MenuID_54       =   53
      MenuText_54     =   "�����Ͽؼ���"
      MenuVisible_54  =   -1  'True
      SubMenuID_54_0  =   0
   End
   Begin �Ͽؼ���UI.DarkWindowBorder DarkWindowBorderSizer 
      Left            =   12120
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      Thickness       =   5
      MinWidth        =   400
      MinHeight       =   75
      Transparency    =   1
      UseSetParent    =   0   'False
   End
   Begin �Ͽؼ���UI.DarkWindowBorder DarkWindowBorder 
      Left            =   11520
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      MinWidth        =   400
      MinHeight       =   75
   End
   Begin �Ͽؼ���UI.DarkTitleBar DarkTitleBar 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13140
      _ExtentX        =   23178
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
      Caption         =   "�¹��� - �Ͽؼ���"
      Picture         =   "frmMain.frx":1BCC2
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   9720
      Top             =   5520
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DockingPane 
      Left            =   10560
      Top             =   5520
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   10
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function LoadSampleFile(FileName As String) As String
    Dim fileHandle As Integer
    Dim FileText As String
    
    If Len(Dir$(FileName$)) = 0 Then
        Err.Raise 53
    End If
    
    fileHandle = FreeFile
    Open FileName$ For Input As #fileHandle
        FileText = Input$(LOF(fileHandle), fileHandle)
    Close #fileHandle
    
    LoadSampleFile = FileText
End Function

Private Sub DockingPane_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            Item.Handle = frmPane.hWnd
            Item.Title = "����"
        
    End Select
End Sub

Private Sub Form_Load()
    Me.DockingPane.AttachToWindow Me.Picture1.hWnd
    Me.DockingPane.CreatePane 1, 100, (Me.ScaleHeight - Me.Picture1.Top) / 15, DockLeftOf
    Me.DockingPane.Options.ShowDockingContextStickers = True
    Me.DockingPane.Options.AlphaDockingContext = True
    Me.DockingPane.Options.LunaColors = False
    Me.DockingPane.Options.ThemedFloatingFrames = True
    
    Me.DarkTitleBar.Caption = DockingPaneGlobalSettings.ResourceImages.LoadFromFile(App.Path & "\Office2010.dll", "Office2010Black.ini")
    Me.DockingPane.VisualTheme = ThemeResource
    Me.DockingPane.PaintManager.SplitterSize = 2
    Me.DockingPane.RedrawPanes
    
    'Me.SyntaxEdit.SetScrollBars Me.HScroll.hWnd, Me.VScroll.hWnd
    Me.SyntaxEdit.PaintManager.BackColor = RGB(28, 28, 28)
    Me.SyntaxEdit.PaintManager.LineNumberBackColor = RGB(28, 28, 28)
    Me.SyntaxEdit.PaintManager.LineNumberTextColor = RGB(86, 156, 214)
    Me.SyntaxEdit.ShowSelectionMargin = False
    Me.SyntaxEdit.DataManager.FileExt = ".cpp"
    Me.SyntaxEdit.SetSyntaxAndColorScheme LoadSampleFile("cpp.schclass"), LoadSampleFile("SyntaxEditDark.ini"), False, False
    
    Me.SkinFramework.LoadSkin "Office2007.cjstyles", "NormalBlue.ini"
    Me.SkinFramework.ApplyWindow Me.hWnd
    
    SetPropA Me.hWnd, "PrevWndProc", SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SetWindowLongA Me.hWnd, GWL_WNDPROC, GetPropA(hWnd, "PrevWndProc")
    Unload frmPane
End Sub

Private Sub Form_Resize()
    Me.Picture1.Height = Me.ScaleHeight - Me.Picture1.Top
    Me.SyntaxEdit.Move 1500, 0, Me.Picture1.ScaleWidth - 3000, Me.Picture1.ScaleHeight - 1500
    Me.picScrollBarCover.ZOrder 0
    Me.picScrollBarCover.Left = Me.SyntaxEdit.Left + Me.SyntaxEdit.Width - Me.picScrollBarCover.Width
    Me.picScrollBarCover.Top = Me.SyntaxEdit.Top + Me.SyntaxEdit.Height - Me.picScrollBarCover.Height
End Sub

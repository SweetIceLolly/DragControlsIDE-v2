VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "CO7FCA~1.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "COE2B7~1.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "�¹��� - �Ͽؼ���"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16845
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Microsoft YaHei UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   16845
   StartUpPosition =   2  'CenterScreen
   Begin DragControlsIDE.DarkMenu DarkMenu 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   495
      Width           =   16845
      _ExtentX        =   29713
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
      MENU_ITEM_COUNT =   66
      LEVELS_COUNT    =   66
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
      LEVELS_40       =   2
      LEVELS_41       =   2
      LEVELS_42       =   2
      LEVELS_43       =   2
      LEVELS_44       =   2
      LEVELS_45       =   2
      LEVELS_46       =   2
      LEVELS_47       =   2
      LEVELS_48       =   2
      LEVELS_49       =   1
      LEVELS_50       =   1
      LEVELS_51       =   1
      LEVELS_52       =   1
      LEVELS_53       =   1
      LEVELS_54       =   1
      LEVELS_55       =   1
      LEVELS_56       =   1
      LEVELS_58       =   1
      LEVELS_59       =   1
      LEVELS_60       =   1
      LEVELS_61       =   1
      LEVELS_62       =   1
      LEVELS_64       =   1
      LEVELS_65       =   1
      LEVELS_66       =   1
      MenuID_1        =   0
      MenuText_1      =   "�ļ�"
      MenuVisible_1   =   -1  'True
      SUBMENU_ITEM_COUNT_1=   6
      SubMenuID_1_0   =   0
      SubMenuText_1_1 =   "�½���Ŀ (&N)       Ctrl+N"
      SubMenuID_1_1   =   2
      SubMenuText_1_2 =   "������Ŀ (&O)       Ctrl+O"
      SubMenuID_1_2   =   3
      SubMenuText_1_3 =   "���� (&S)           Ctrl+S"
      SubMenuID_1_3   =   4
      SubMenuText_1_4 =   "���Ϊ (&A)         Ctrl+Shift+S"
      SubMenuID_1_4   =   5
      SubMenuText_1_5 =   "-"
      SubMenuID_1_5   =   6
      SubMenuText_1_6 =   "�˳� (&E)"
      SubMenuID_1_6   =   7
      MenuID_2        =   1
      MenuText_2      =   "�½���Ŀ (&N)       Ctrl+N"
      MenuVisible_2   =   -1  'True
      SubMenuID_2_0   =   0
      MenuID_3        =   2
      MenuText_3      =   "������Ŀ (&O)       Ctrl+O"
      MenuVisible_3   =   -1  'True
      SubMenuID_3_0   =   0
      MenuID_4        =   3
      MenuText_4      =   "���� (&S)           Ctrl+S"
      MenuVisible_4   =   -1  'True
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "���Ϊ (&A)         Ctrl+Shift+S"
      MenuVisible_5   =   -1  'True
      SubMenuID_5_0   =   0
      MenuID_6        =   5
      MenuText_6      =   "-"
      MenuVisible_6   =   -1  'True
      SubMenuID_6_0   =   0
      MenuID_7        =   6
      MenuText_7      =   "�˳� (&E)"
      MenuVisible_7   =   -1  'True
      SubMenuID_7_0   =   0
      MenuID_8        =   7
      MenuText_8      =   "�༭"
      MenuVisible_8   =   -1  'True
      SUBMENU_ITEM_COUNT_8=   19
      SubMenuID_8_0   =   0
      SubMenuText_8_1 =   "���� (&U)           Ctrl+Z"
      SubMenuID_8_1   =   9
      SubMenuText_8_2 =   "�ظ� (&Y)           Ctrl+Y"
      SubMenuID_8_2   =   10
      SubMenuText_8_3 =   "-"
      SubMenuID_8_3   =   11
      SubMenuText_8_4 =   "���� (&U)           Ctrl+X"
      SubMenuID_8_4   =   12
      SubMenuText_8_5 =   "���� (&C)           Ctrl+C"
      SubMenuID_8_5   =   13
      SubMenuText_8_6 =   "ճ�� (&P)           Ctrl+V"
      SubMenuID_8_6   =   14
      SubMenuText_8_7 =   "ȫѡ (&S)           Ctrl+A"
      SubMenuID_8_7   =   15
      SubMenuText_8_8 =   "ɾ���� (&D)         Ctrl+L"
      SubMenuID_8_8   =   16
      SubMenuText_8_9 =   "-"
      SubMenuID_8_9   =   17
      SubMenuText_8_10=   "���� (&F)           Ctrl+F"
      SubMenuID_8_10  =   18
      SubMenuText_8_11=   "�滻 (&R)           Ctrl+H"
      SubMenuID_8_11  =   19
      SubMenuText_8_12=   "-"
      SubMenuID_8_12  =   20
      SubMenuText_8_13=   "�������� (&I)       Tab"
      SubMenuID_8_13  =   21
      SubMenuText_8_14=   "�������� (&O)       Shift+Tab"
      SubMenuID_8_14  =   22
      SubMenuText_8_15=   "-"
      SubMenuID_8_15  =   23
      SubMenuText_8_16=   "���/�Ƴ��ϵ� (&B)  F9"
      SubMenuID_8_16  =   24
      SubMenuText_8_17=   "������жϵ� (&M)"
      SubMenuID_8_17  =   25
      SubMenuText_8_18=   "-"
      SubMenuID_8_18  =   26
      SubMenuText_8_19=   "��ת���� (&J)       Ctrl+G"
      SubMenuID_8_19  =   27
      MenuID_9        =   8
      MenuText_9      =   "���� (&U)           Ctrl+Z"
      MenuVisible_9   =   -1  'True
      SubMenuID_9_0   =   0
      MenuID_10       =   9
      MenuText_10     =   "�ظ� (&Y)           Ctrl+Y"
      MenuVisible_10  =   -1  'True
      SubMenuID_10_0  =   0
      MenuID_11       =   10
      MenuText_11     =   "-"
      MenuVisible_11  =   -1  'True
      SubMenuID_11_0  =   0
      MenuID_12       =   11
      MenuText_12     =   "���� (&U)           Ctrl+X"
      MenuVisible_12  =   -1  'True
      SubMenuID_12_0  =   0
      MenuID_13       =   12
      MenuText_13     =   "���� (&C)           Ctrl+C"
      MenuVisible_13  =   -1  'True
      SubMenuID_13_0  =   0
      MenuID_14       =   13
      MenuText_14     =   "ճ�� (&P)           Ctrl+V"
      MenuVisible_14  =   -1  'True
      SubMenuID_14_0  =   0
      MenuID_15       =   14
      MenuText_15     =   "ȫѡ (&S)           Ctrl+A"
      MenuVisible_15  =   -1  'True
      SubMenuID_15_0  =   0
      MenuID_16       =   15
      MenuText_16     =   "ɾ���� (&D)         Ctrl+L"
      MenuVisible_16  =   -1  'True
      SubMenuID_16_0  =   0
      MenuID_17       =   16
      MenuText_17     =   "-"
      MenuVisible_17  =   -1  'True
      SubMenuID_17_0  =   0
      MenuID_18       =   17
      MenuText_18     =   "���� (&F)           Ctrl+F"
      MenuVisible_18  =   -1  'True
      SubMenuID_18_0  =   0
      MenuID_19       =   18
      MenuText_19     =   "�滻 (&R)           Ctrl+H"
      MenuVisible_19  =   -1  'True
      SubMenuID_19_0  =   0
      MenuID_20       =   19
      MenuText_20     =   "-"
      MenuVisible_20  =   -1  'True
      SubMenuID_20_0  =   0
      MenuID_21       =   20
      MenuText_21     =   "�������� (&I)       Tab"
      MenuVisible_21  =   -1  'True
      SubMenuID_21_0  =   0
      MenuID_22       =   21
      MenuText_22     =   "�������� (&O)       Shift+Tab"
      MenuVisible_22  =   -1  'True
      SubMenuID_22_0  =   0
      MenuID_23       =   22
      MenuText_23     =   "-"
      MenuVisible_23  =   -1  'True
      SubMenuID_23_0  =   0
      MenuID_24       =   23
      MenuText_24     =   "���/�Ƴ��ϵ� (&B)  F9"
      MenuVisible_24  =   -1  'True
      SubMenuID_24_0  =   0
      MenuID_25       =   24
      MenuText_25     =   "������жϵ� (&M)"
      MenuVisible_25  =   -1  'True
      SubMenuID_25_0  =   0
      MenuID_26       =   25
      MenuText_26     =   "-"
      MenuVisible_26  =   -1  'True
      SubMenuID_26_0  =   0
      MenuID_27       =   26
      MenuText_27     =   "��ת���� (&J)       Ctrl+G"
      MenuVisible_27  =   -1  'True
      SubMenuID_27_0  =   0
      MenuID_28       =   27
      MenuText_28     =   "��ͼ"
      MenuVisible_28  =   -1  'True
      SUBMENU_ITEM_COUNT_28=   6
      SubMenuID_28_0  =   0
      SubMenuText_28_1=   "������ (&T)"
      SubMenuID_28_1  =   29
      SubMenuText_28_2=   "�ؼ��� (&C)"
      SubMenuID_28_2  =   30
      SubMenuText_28_3=   "���� (&P)           F4"
      SubMenuID_28_3  =   31
      SubMenuText_28_4=   "������Դ������ (&M)"
      SubMenuID_28_4  =   32
      SubMenuText_28_5=   "�����б� (&E)       Ctrl+E"
      SubMenuID_28_5  =   33
      SubMenuText_28_6=   "��� (&O)           Ctrl+Alt+O"
      SubMenuID_28_6  =   34
      MenuID_29       =   28
      MenuText_29     =   "������ (&T)"
      MenuVisible_29  =   -1  'True
      SubMenuID_29_0  =   0
      MenuID_30       =   29
      MenuText_30     =   "�ؼ��� (&C)"
      MenuVisible_30  =   -1  'True
      SubMenuID_30_0  =   0
      MenuID_31       =   30
      MenuText_31     =   "���� (&P)           F4"
      MenuVisible_31  =   -1  'True
      SubMenuID_31_0  =   0
      MenuID_32       =   31
      MenuText_32     =   "������Դ������ (&M)"
      MenuVisible_32  =   -1  'True
      SubMenuID_32_0  =   0
      MenuID_33       =   32
      MenuText_33     =   "�����б� (&E)       Ctrl+E"
      MenuVisible_33  =   -1  'True
      SubMenuID_33_0  =   0
      MenuID_34       =   33
      MenuText_34     =   "��� (&O)           Ctrl+Alt+O"
      MenuVisible_34  =   -1  'True
      SubMenuID_34_0  =   0
      MenuID_35       =   34
      MenuText_35     =   "����"
      MenuVisible_35  =   -1  'True
      SUBMENU_ITEM_COUNT_35=   2
      SubMenuID_35_0  =   0
      SubMenuText_35_1=   "���ɴ����ļ� (&C)"
      SubMenuID_35_1  =   36
      SubMenuText_35_2=   "���ɿ�ִ���ļ� (&E) Ctrl+F5"
      SubMenuID_35_2  =   37
      MenuID_36       =   35
      MenuText_36     =   "���ɴ����ļ� (&C)"
      MenuVisible_36  =   -1  'True
      SubMenuID_36_0  =   0
      MenuID_37       =   36
      MenuText_37     =   "���ɿ�ִ���ļ� (&E) Ctrl+F5"
      MenuVisible_37  =   -1  'True
      SubMenuID_37_0  =   0
      MenuID_38       =   37
      MenuText_38     =   "����"
      MenuVisible_38  =   -1  'True
      SUBMENU_ITEM_COUNT_38=   9
      SubMenuID_38_0  =   0
      SubMenuText_38_1=   "����"
      SubMenuID_38_1  =   39
      SubMenuText_38_2=   "���� (&R)           F5"
      SubMenuID_38_2  =   49
      SubMenuText_38_3=   "�ж� (&B)           Ctrl+Alt+Break"
      SubMenuID_38_3  =   50
      SubMenuText_38_4=   "ֹͣ (&E)           Shift+F5"
      SubMenuID_38_4  =   51
      SubMenuText_38_5=   "�������� (&S)       Ctrl+Shift+F5"
      SubMenuID_38_5  =   52
      SubMenuText_38_6=   "-"
      SubMenuID_38_6  =   53
      SubMenuText_38_7=   "�����ִ��         F11"
      SubMenuID_38_7  =   54
      SubMenuText_38_8=   "�����ִ��         F10"
      SubMenuID_38_8  =   55
      SubMenuText_38_9=   "ִ�е�����         Shift+F11"
      SubMenuID_38_9  =   56
      MenuID_39       =   38
      MenuText_39     =   "����"
      MenuVisible_39  =   -1  'True
      SUBMENU_ITEM_COUNT_39=   9
      SubMenuID_39_0  =   0
      SubMenuText_39_1=   "�ϵ��б� (&B)       Ctrl+Alt+B"
      SubMenuID_39_1  =   40
      SubMenuText_39_2=   "���Ӵ��� (&W)       Ctrl+Alt+W"
      SubMenuID_39_2  =   41
      SubMenuText_39_3=   "���� (&L)           Ctrl+Alt+L"
      SubMenuID_39_3  =   42
      SubMenuText_39_4=   "���ö�ջ (&C)       Ctrl+Alt+C"
      SubMenuID_39_4  =   43
      SubMenuText_39_5=   "�߳� (&T)           Ctrl+Alt+T"
      SubMenuID_39_5  =   44
      SubMenuText_39_6=   "ģ�� (&M)           Ctrl+Alt+M"
      SubMenuID_39_6  =   45
      SubMenuText_39_7=   "�ڴ� (&E)           Ctrl+Alt+E"
      SubMenuID_39_7  =   46
      SubMenuText_39_8=   "�Ĵ��� (&R)         Ctrl+Alt+R"
      SubMenuID_39_8  =   47
      SubMenuText_39_9=   "����� (&D)         Ctrl+Alt+D"
      SubMenuID_39_9  =   48
      MenuID_40       =   39
      MenuText_40     =   "�ϵ��б� (&B)       Ctrl+Alt+B"
      MenuVisible_40  =   -1  'True
      SubMenuID_40_0  =   0
      MenuID_41       =   40
      MenuText_41     =   "���Ӵ��� (&W)       Ctrl+Alt+W"
      MenuVisible_41  =   -1  'True
      SubMenuID_41_0  =   0
      MenuID_42       =   41
      MenuText_42     =   "���� (&L)           Ctrl+Alt+L"
      MenuVisible_42  =   -1  'True
      SubMenuID_42_0  =   0
      MenuID_43       =   42
      MenuText_43     =   "���ö�ջ (&C)       Ctrl+Alt+C"
      MenuVisible_43  =   -1  'True
      SubMenuID_43_0  =   0
      MenuID_44       =   43
      MenuText_44     =   "�߳� (&T)           Ctrl+Alt+T"
      MenuVisible_44  =   -1  'True
      SubMenuID_44_0  =   0
      MenuID_45       =   44
      MenuText_45     =   "ģ�� (&M)           Ctrl+Alt+M"
      MenuVisible_45  =   -1  'True
      SubMenuID_45_0  =   0
      MenuID_46       =   45
      MenuText_46     =   "�ڴ� (&E)           Ctrl+Alt+E"
      MenuVisible_46  =   -1  'True
      SubMenuID_46_0  =   0
      MenuID_47       =   46
      MenuText_47     =   "�Ĵ��� (&R)         Ctrl+Alt+R"
      MenuVisible_47  =   -1  'True
      SubMenuID_47_0  =   0
      MenuID_48       =   47
      MenuText_48     =   "����� (&D)         Ctrl+Alt+D"
      MenuVisible_48  =   -1  'True
      SubMenuID_48_0  =   0
      MenuID_49       =   48
      MenuText_49     =   "���� (&R)           F5"
      MenuVisible_49  =   -1  'True
      SubMenuID_49_0  =   0
      MenuID_50       =   49
      MenuText_50     =   "�ж� (&B)           Ctrl+Alt+Break"
      MenuVisible_50  =   -1  'True
      SubMenuID_50_0  =   0
      MenuID_51       =   50
      MenuText_51     =   "ֹͣ (&E)           Shift+F5"
      MenuVisible_51  =   -1  'True
      SubMenuID_51_0  =   0
      MenuID_52       =   51
      MenuText_52     =   "�������� (&S)       Ctrl+Shift+F5"
      MenuVisible_52  =   -1  'True
      SubMenuID_52_0  =   0
      MenuID_53       =   52
      MenuText_53     =   "-"
      MenuVisible_53  =   -1  'True
      SubMenuID_53_0  =   0
      MenuID_54       =   53
      MenuText_54     =   "�����ִ��         F11"
      MenuVisible_54  =   -1  'True
      SubMenuID_54_0  =   0
      MenuID_55       =   54
      MenuText_55     =   "�����ִ��         F10"
      MenuVisible_55  =   -1  'True
      SubMenuID_55_0  =   0
      MenuID_56       =   55
      MenuText_56     =   "ִ�е�����         Shift+F11"
      MenuVisible_56  =   -1  'True
      SubMenuID_56_0  =   0
      MenuID_57       =   56
      MenuText_57     =   "����"
      MenuVisible_57  =   -1  'True
      SUBMENU_ITEM_COUNT_57=   5
      SubMenuID_57_0  =   0
      SubMenuText_57_1=   "���ڹ��� (&W)"
      SubMenuID_57_1  =   58
      SubMenuText_57_2=   "��Ϣ���� (&M)"
      SubMenuID_57_2  =   59
      SubMenuText_57_3=   "���� (&P)"
      SubMenuID_57_3  =   60
      SubMenuText_57_4=   "-"
      SubMenuID_57_4  =   61
      SubMenuText_57_5=   "���� (&O)"
      SubMenuID_57_5  =   62
      MenuID_58       =   57
      MenuText_58     =   "���ڹ��� (&W)"
      MenuVisible_58  =   -1  'True
      SubMenuID_58_0  =   0
      MenuID_59       =   58
      MenuText_59     =   "��Ϣ���� (&M)"
      MenuVisible_59  =   -1  'True
      SubMenuID_59_0  =   0
      MenuID_60       =   59
      MenuText_60     =   "���� (&P)"
      MenuVisible_60  =   -1  'True
      SubMenuID_60_0  =   0
      MenuID_61       =   60
      MenuText_61     =   "-"
      MenuVisible_61  =   -1  'True
      SubMenuID_61_0  =   0
      MenuID_62       =   61
      MenuText_62     =   "���� (&O)"
      MenuVisible_62  =   -1  'True
      SubMenuID_62_0  =   0
      MenuID_63       =   62
      MenuText_63     =   "����"
      MenuVisible_63  =   -1  'True
      SUBMENU_ITEM_COUNT_63=   3
      SubMenuID_63_0  =   0
      SubMenuText_63_1=   "�����ĵ� (&D)       F1"
      SubMenuID_63_1  =   64
      SubMenuText_63_2=   "ʾ������ (&E)"
      SubMenuID_63_2  =   65
      SubMenuText_63_3=   "�����Ͽؼ��� (&A) Ctrl+F1"
      SubMenuID_63_3  =   66
      MenuID_64       =   63
      MenuText_64     =   "�����ĵ� (&D)       F1"
      MenuVisible_64  =   -1  'True
      SubMenuID_64_0  =   0
      MenuID_65       =   64
      MenuText_65     =   "ʾ������ (&E)"
      MenuVisible_65  =   -1  'True
      SubMenuID_65_0  =   0
      MenuID_66       =   65
      MenuText_66     =   "�����Ͽؼ��� (&A) Ctrl+F1"
      MenuVisible_66  =   -1  'True
      SubMenuID_66_0  =   0
   End
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00302D2D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   16845
      TabIndex        =   2
      Top             =   810
      Width           =   16845
   End
   Begin VB.PictureBox picClientArea 
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
      ScaleWidth      =   16845
      TabIndex        =   1
      Top             =   1185
      Width           =   16845
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorderSizer 
      Left            =   12120
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      Thickness       =   3
      MinWidth        =   400
      MinHeight       =   100
      Transparency    =   1
      UseSetParent    =   0   'False
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   11520
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      MinWidth        =   400
      MinHeight       =   100
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16845
      _ExtentX        =   29713
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
      BindCaption     =   -1  'True
      Picture         =   "frmMain.frx":1BCC2
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   10080
      Top             =   5520
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DockingPane 
      Left            =   10800
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
'====================================================
'����:      ������
'����:      ����
'�ļ�:      frmMain.frm
'====================================================

Option Explicit

'��ȡ���������С��״̬
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Private Sub DockingPane_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            
        
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    '�����ַ�����Դ
    
    
    '����������
    Me.DockingPane.AttachToWindow Me.picClientArea.hWnd                                                                 '�󶨹�����
    
    '����Docking Pane����ʽ
    Me.DockingPane.Options.ShowDockingContextStickers = True                                                            '��ʾDocking stickers
    Me.DockingPane.Options.AlphaDockingContext = True                                                                   '�ƶ���ʱ��͸��
    Me.DockingPane.Options.ThemedFloatingFrames = True                                                                  '��Ϊ����ʱ�߿򱣳���ʽ
    Me.DockingPane.Options.ShowContentsWhileDragging = True
    If DockingPaneGlobalSettings.ResourceImages.LoadFromFile(GetAppPath & "Skin.dll", "Office2010Black.ini") = False Then
        MsgBox "������ʽʧ�ܣ�", vbCritical, "����"
    End If
    Me.DockingPane.VisualTheme = ThemeResource                                                                          '����Ϊ����Դ�ļ���ȡ��ʽ
    Me.DockingPane.PaintManager.SplitterSize = 2                                                                        '���÷ָ�����Ĵ�С
    
    '����Ƥ��
    Me.SkinFramework.RemoveAllWindows
    Me.SkinFramework.LoadSkin "Skin.cjstyles", "NormalBlue.ini"
    Me.SkinFramework.ApplyWindow Me.hWnd

    '���ô������໯�������������
    SetPropA Me.hWnd, "PrevWndProc", SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf MainWindowMaximizeCloseFixProc)
    
    frmCodeWindow.Show
    SetPropA frmCodeWindow.hWnd, "Parent", Me.picClientArea.hWnd
    SetParent frmCodeWindow.hWnd, Me.picClientArea.hWnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '�ָ��������໯
    SetWindowLongA Me.hWnd, GWL_WNDPROC, GetPropA(hWnd, "PrevWndProc")
    Unload frmCodeWindow
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    '�������ͻ������Ĵ�С
    Me.picClientArea.Height = Me.ScaleHeight - Me.picClientArea.Top
    
    '������󻯵��Ӵ���Ĵ�С
    Dim Target  As Form
    Dim wp      As WINDOWPLACEMENT
    
    For Each Target In Forms
        If GetPropA(Target.hWnd, "Parent") = Me.picClientArea.hWnd Then
            GetWindowPlacement Target.hWnd, wp
            If wp.ShowCmd = SW_MAXIMIZE Then
                ShowWindow Target.hWnd, SW_HIDE
                ShowWindow Target.hWnd, SW_MAXIMIZE
            End If
        End If
    Next Target
End Sub

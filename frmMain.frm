VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "DockingPane.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "SkinFramework.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "�Ͽؼ���"
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
   Begin VB.Timer tmrCheckProcess 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   13440
      Top             =   7320
   End
   Begin DragControlsIDE.DarkMenu DarkMenu 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   495
      Width           =   16845
      _ExtentX        =   29713
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MENU_ITEM_COUNT =   70
      LEVELS_COUNT    =   70
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
      LEVELS_49       =   2
      LEVELS_50       =   2
      LEVELS_51       =   2
      LEVELS_52       =   2
      LEVELS_53       =   1
      LEVELS_54       =   1
      LEVELS_55       =   1
      LEVELS_56       =   1
      LEVELS_57       =   1
      LEVELS_58       =   1
      LEVELS_59       =   1
      LEVELS_60       =   1
      LEVELS_62       =   1
      LEVELS_63       =   1
      LEVELS_64       =   1
      LEVELS_65       =   1
      LEVELS_66       =   1
      LEVELS_68       =   1
      LEVELS_69       =   1
      LEVELS_70       =   1
      MenuID_1        =   0
      MenuText_1      =   "�ļ�"
      MenuVisible_1   =   -1  'True
      MenuIcon_1      =   "frmMain.frx":1BCC2
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
      MenuIcon_2      =   "frmMain.frx":1BCE2
      SubMenuID_2_0   =   0
      MenuID_3        =   2
      MenuText_3      =   "������Ŀ (&O)       Ctrl+O"
      MenuVisible_3   =   -1  'True
      MenuIcon_3      =   "frmMain.frx":1BE11
      SubMenuID_3_0   =   0
      MenuID_4        =   3
      MenuText_4      =   "���� (&S)           Ctrl+S"
      MenuVisible_4   =   -1  'True
      MenuIcon_4      =   "frmMain.frx":1BFFB
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "���Ϊ (&A)         Ctrl+Shift+S"
      MenuVisible_5   =   -1  'True
      MenuIcon_5      =   "frmMain.frx":1C109
      SubMenuID_5_0   =   0
      MenuID_6        =   5
      MenuText_6      =   "-"
      MenuVisible_6   =   -1  'True
      MenuIcon_6      =   "frmMain.frx":1C2C8
      SubMenuID_6_0   =   0
      MenuID_7        =   6
      MenuText_7      =   "�˳� (&E)"
      MenuVisible_7   =   -1  'True
      MenuIcon_7      =   "frmMain.frx":1C2E8
      SubMenuID_7_0   =   0
      MenuID_8        =   7
      MenuText_8      =   "�༭"
      MenuVisible_8   =   -1  'True
      MenuIcon_8      =   "frmMain.frx":1C308
      SUBMENU_ITEM_COUNT_8=   19
      SubMenuID_8_0   =   0
      SubMenuText_8_1 =   "���� (&U)           Ctrl+Z"
      SubMenuID_8_1   =   9
      SubMenuText_8_2 =   "�ظ� (&R)           Ctrl+Y"
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
      SubMenuText_8_11=   "�滻 (&E)           Ctrl+H"
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
      MenuIcon_9      =   "frmMain.frx":1C328
      SubMenuID_9_0   =   0
      MenuID_10       =   9
      MenuText_10     =   "�ظ� (&R)           Ctrl+Y"
      MenuVisible_10  =   -1  'True
      MenuIcon_10     =   "frmMain.frx":1C4F9
      SubMenuID_10_0  =   0
      MenuID_11       =   10
      MenuText_11     =   "-"
      MenuVisible_11  =   -1  'True
      MenuIcon_11     =   "frmMain.frx":1C708
      SubMenuID_11_0  =   0
      MenuID_12       =   11
      MenuText_12     =   "���� (&U)           Ctrl+X"
      MenuVisible_12  =   -1  'True
      MenuIcon_12     =   "frmMain.frx":1C728
      SubMenuID_12_0  =   0
      MenuID_13       =   12
      MenuText_13     =   "���� (&C)           Ctrl+C"
      MenuVisible_13  =   -1  'True
      MenuIcon_13     =   "frmMain.frx":1C888
      SubMenuID_13_0  =   0
      MenuID_14       =   13
      MenuText_14     =   "ճ�� (&P)           Ctrl+V"
      MenuVisible_14  =   -1  'True
      MenuIcon_14     =   "frmMain.frx":1C9E3
      SubMenuID_14_0  =   0
      MenuID_15       =   14
      MenuText_15     =   "ȫѡ (&S)           Ctrl+A"
      MenuVisible_15  =   -1  'True
      MenuIcon_15     =   "frmMain.frx":1CB3B
      SubMenuID_15_0  =   0
      MenuID_16       =   15
      MenuText_16     =   "ɾ���� (&D)         Ctrl+L"
      MenuVisible_16  =   -1  'True
      MenuIcon_16     =   "frmMain.frx":1CC3A
      SubMenuID_16_0  =   0
      MenuID_17       =   16
      MenuText_17     =   "-"
      MenuVisible_17  =   -1  'True
      MenuIcon_17     =   "frmMain.frx":1CC5A
      SubMenuID_17_0  =   0
      MenuID_18       =   17
      MenuText_18     =   "���� (&F)           Ctrl+F"
      MenuVisible_18  =   -1  'True
      MenuIcon_18     =   "frmMain.frx":1CC7A
      SubMenuID_18_0  =   0
      MenuID_19       =   18
      MenuText_19     =   "�滻 (&E)           Ctrl+H"
      MenuVisible_19  =   -1  'True
      MenuIcon_19     =   "frmMain.frx":1CD8D
      SubMenuID_19_0  =   0
      MenuID_20       =   19
      MenuText_20     =   "-"
      MenuVisible_20  =   -1  'True
      MenuIcon_20     =   "frmMain.frx":1CEF1
      SubMenuID_20_0  =   0
      MenuID_21       =   20
      MenuText_21     =   "�������� (&I)       Tab"
      MenuVisible_21  =   -1  'True
      MenuIcon_21     =   "frmMain.frx":1CF11
      SubMenuID_21_0  =   0
      MenuID_22       =   21
      MenuText_22     =   "�������� (&O)       Shift+Tab"
      MenuVisible_22  =   -1  'True
      MenuIcon_22     =   "frmMain.frx":1D268
      SubMenuID_22_0  =   0
      MenuID_23       =   22
      MenuText_23     =   "-"
      MenuVisible_23  =   -1  'True
      MenuIcon_23     =   "frmMain.frx":1D5BF
      SubMenuID_23_0  =   0
      MenuID_24       =   23
      MenuText_24     =   "���/�Ƴ��ϵ� (&B)  F9"
      MenuVisible_24  =   -1  'True
      MenuIcon_24     =   "frmMain.frx":1D5DF
      SubMenuID_24_0  =   0
      MenuID_25       =   24
      MenuText_25     =   "������жϵ� (&M)"
      MenuVisible_25  =   -1  'True
      MenuIcon_25     =   "frmMain.frx":1D5FF
      SubMenuID_25_0  =   0
      MenuID_26       =   25
      MenuText_26     =   "-"
      MenuVisible_26  =   -1  'True
      MenuIcon_26     =   "frmMain.frx":1D61F
      SubMenuID_26_0  =   0
      MenuID_27       =   26
      MenuText_27     =   "��ת���� (&J)       Ctrl+G"
      MenuVisible_27  =   -1  'True
      MenuIcon_27     =   "frmMain.frx":1D63F
      SubMenuID_27_0  =   0
      MenuID_28       =   27
      MenuText_28     =   "��ͼ"
      MenuVisible_28  =   -1  'True
      MenuIcon_28     =   "frmMain.frx":1D65F
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
      MenuCheckBox_29 =   -1  'True
      MenuVisible_29  =   -1  'True
      MenuIcon_29     =   "frmMain.frx":1D67F
      SubMenuID_29_0  =   0
      MenuID_30       =   29
      MenuText_30     =   "�ؼ��� (&C)"
      MenuCheckBox_30 =   -1  'True
      MenuVisible_30  =   -1  'True
      MenuIcon_30     =   "frmMain.frx":1D744
      SubMenuID_30_0  =   0
      MenuID_31       =   30
      MenuText_31     =   "���� (&P)           F4"
      MenuCheckBox_31 =   -1  'True
      MenuVisible_31  =   -1  'True
      MenuIcon_31     =   "frmMain.frx":1D7F7
      SubMenuID_31_0  =   0
      MenuID_32       =   31
      MenuText_32     =   "������Դ������ (&M)"
      MenuCheckBox_32 =   -1  'True
      MenuVisible_32  =   -1  'True
      MenuIcon_32     =   "frmMain.frx":1D983
      SubMenuID_32_0  =   0
      MenuID_33       =   32
      MenuText_33     =   "�����б� (&E)       Ctrl+E"
      MenuCheckBox_33 =   -1  'True
      MenuVisible_33  =   -1  'True
      MenuIcon_33     =   "frmMain.frx":1D9A3
      SubMenuID_33_0  =   0
      MenuID_34       =   33
      MenuText_34     =   "��� (&O)           Ctrl+Alt+O"
      MenuCheckBox_34 =   -1  'True
      MenuVisible_34  =   -1  'True
      MenuIcon_34     =   "frmMain.frx":1DB51
      SubMenuID_34_0  =   0
      MenuID_35       =   34
      MenuText_35     =   "����"
      MenuVisible_35  =   -1  'True
      MenuIcon_35     =   "frmMain.frx":1DCB9
      SUBMENU_ITEM_COUNT_35=   2
      SubMenuID_35_0  =   0
      SubMenuText_35_1=   "���ɴ����ļ� (&C)"
      SubMenuID_35_1  =   36
      SubMenuText_35_2=   "���ɿ�ִ���ļ� (&E) Ctrl+F5"
      SubMenuID_35_2  =   37
      MenuID_36       =   35
      MenuText_36     =   "���ɴ����ļ� (&C)"
      MenuVisible_36  =   -1  'True
      MenuIcon_36     =   "frmMain.frx":1DCD9
      SubMenuID_36_0  =   0
      MenuID_37       =   36
      MenuText_37     =   "���ɿ�ִ���ļ� (&E) Ctrl+F5"
      MenuVisible_37  =   -1  'True
      MenuIcon_37     =   "frmMain.frx":1DEE9
      SubMenuID_37_0  =   0
      MenuID_38       =   37
      MenuText_38     =   "����"
      MenuVisible_38  =   -1  'True
      MenuIcon_38     =   "frmMain.frx":1DF09
      SUBMENU_ITEM_COUNT_38=   9
      SubMenuID_38_0  =   0
      SubMenuText_38_1=   "����"
      SubMenuID_38_1  =   39
      SubMenuText_38_2=   "���� (&R)           F5"
      SubMenuID_38_2  =   53
      SubMenuText_38_3=   "�ж� (&B)           Ctrl+Alt+Break"
      SubMenuID_38_3  =   54
      SubMenuText_38_4=   "ֹͣ (&E)           Shift+F5"
      SubMenuID_38_4  =   55
      SubMenuText_38_5=   "�������� (&S)       Ctrl+Shift+F5"
      SubMenuID_38_5  =   56
      SubMenuText_38_6=   "-"
      SubMenuID_38_6  =   57
      SubMenuText_38_7=   "�����ִ��         F11"
      SubMenuID_38_7  =   58
      SubMenuText_38_8=   "�����ִ��         F10"
      SubMenuID_38_8  =   59
      SubMenuText_38_9=   "ִ�е�����         Shift+F11"
      SubMenuID_38_9  =   60
      MenuID_39       =   38
      MenuText_39     =   "����"
      MenuVisible_39  =   -1  'True
      MenuIcon_39     =   "frmMain.frx":1DF29
      SUBMENU_ITEM_COUNT_39=   13
      SubMenuID_39_0  =   0
      SubMenuText_39_1=   "�ϵ��б� (&B)       Ctrl+Alt+B"
      SubMenuID_39_1  =   40
      SubMenuText_39_2=   "-"
      SubMenuID_39_2  =   41
      SubMenuText_39_3=   "���Ӵ��� (&W)       Ctrl+Alt+W"
      SubMenuID_39_3  =   42
      SubMenuText_39_4=   "���� (&L)           Ctrl+Alt+L"
      SubMenuID_39_4  =   43
      SubMenuText_39_5=   "�������� (&I)       Ctrl+Alt+I"
      SubMenuID_39_5  =   44
      SubMenuText_39_6=   "-"
      SubMenuID_39_6  =   45
      SubMenuText_39_7=   "���ö�ջ (&C)       Ctrl+Alt+C"
      SubMenuID_39_7  =   46
      SubMenuText_39_8=   "�߳� (&T)           Ctrl+Alt+T"
      SubMenuID_39_8  =   47
      SubMenuText_39_9=   "ģ�� (&M)           Ctrl+Alt+M"
      SubMenuID_39_9  =   48
      SubMenuText_39_10=   "-"
      SubMenuID_39_10 =   49
      SubMenuText_39_11=   "�ڴ� (&E)           Ctrl+Alt+E"
      SubMenuID_39_11 =   50
      SubMenuText_39_12=   "�Ĵ��� (&R)         Ctrl+Alt+R"
      SubMenuID_39_12 =   51
      SubMenuText_39_13=   "����� (&D)         Ctrl+Alt+D"
      SubMenuID_39_13 =   52
      MenuID_40       =   39
      MenuText_40     =   "�ϵ��б� (&B)       Ctrl+Alt+B"
      MenuCheckBox_40 =   -1  'True
      MenuVisible_40  =   -1  'True
      MenuIcon_40     =   "frmMain.frx":1DF49
      SubMenuID_40_0  =   0
      MenuID_41       =   40
      MenuText_41     =   "-"
      MenuVisible_41  =   -1  'True
      MenuIcon_41     =   "frmMain.frx":1E050
      SubMenuID_41_0  =   0
      MenuID_42       =   41
      MenuText_42     =   "���Ӵ��� (&W)       Ctrl+Alt+W"
      MenuCheckBox_42 =   -1  'True
      MenuVisible_42  =   -1  'True
      MenuIcon_42     =   "frmMain.frx":1E070
      SubMenuID_42_0  =   0
      MenuID_43       =   42
      MenuText_43     =   "���� (&L)           Ctrl+Alt+L"
      MenuCheckBox_43 =   -1  'True
      MenuVisible_43  =   -1  'True
      MenuIcon_43     =   "frmMain.frx":1E1A7
      SubMenuID_43_0  =   0
      MenuID_44       =   43
      MenuText_44     =   "�������� (&I)       Ctrl+Alt+I"
      MenuCheckBox_44 =   -1  'True
      MenuVisible_44  =   -1  'True
      MenuIcon_44     =   "frmMain.frx":1E1C7
      SubMenuID_44_0  =   0
      MenuID_45       =   44
      MenuText_45     =   "-"
      MenuVisible_45  =   -1  'True
      MenuIcon_45     =   "frmMain.frx":1E1E7
      SubMenuID_45_0  =   0
      MenuID_46       =   45
      MenuText_46     =   "���ö�ջ (&C)       Ctrl+Alt+C"
      MenuCheckBox_46 =   -1  'True
      MenuVisible_46  =   -1  'True
      MenuIcon_46     =   "frmMain.frx":1E207
      SubMenuID_46_0  =   0
      MenuID_47       =   46
      MenuText_47     =   "�߳� (&T)           Ctrl+Alt+T"
      MenuCheckBox_47 =   -1  'True
      MenuVisible_47  =   -1  'True
      MenuIcon_47     =   "frmMain.frx":1E227
      SubMenuID_47_0  =   0
      MenuID_48       =   47
      MenuText_48     =   "ģ�� (&M)           Ctrl+Alt+M"
      MenuCheckBox_48 =   -1  'True
      MenuVisible_48  =   -1  'True
      MenuIcon_48     =   "frmMain.frx":1E4A0
      SubMenuID_48_0  =   0
      MenuID_49       =   48
      MenuText_49     =   "-"
      MenuVisible_49  =   -1  'True
      MenuIcon_49     =   "frmMain.frx":1E564
      SubMenuID_49_0  =   0
      MenuID_50       =   49
      MenuText_50     =   "�ڴ� (&E)           Ctrl+Alt+E"
      MenuCheckBox_50 =   -1  'True
      MenuVisible_50  =   -1  'True
      MenuIcon_50     =   "frmMain.frx":1E584
      SubMenuID_50_0  =   0
      MenuID_51       =   50
      MenuText_51     =   "�Ĵ��� (&R)         Ctrl+Alt+R"
      MenuCheckBox_51 =   -1  'True
      MenuVisible_51  =   -1  'True
      MenuIcon_51     =   "frmMain.frx":1E667
      SubMenuID_51_0  =   0
      MenuID_52       =   51
      MenuText_52     =   "����� (&D)         Ctrl+Alt+D"
      MenuCheckBox_52 =   -1  'True
      MenuVisible_52  =   -1  'True
      MenuIcon_52     =   "frmMain.frx":1E687
      SubMenuID_52_0  =   0
      MenuID_53       =   52
      MenuText_53     =   "���� (&R)           F5"
      MenuVisible_53  =   -1  'True
      MenuIcon_53     =   "frmMain.frx":1E73D
      SubMenuID_53_0  =   0
      MenuID_54       =   53
      MenuText_54     =   "�ж� (&B)           Ctrl+Alt+Break"
      MenuVisible_54  =   -1  'True
      MenuIcon_54     =   "frmMain.frx":1E9C2
      SubMenuID_54_0  =   0
      MenuID_55       =   54
      MenuText_55     =   "ֹͣ (&E)           Shift+F5"
      MenuVisible_55  =   -1  'True
      MenuIcon_55     =   "frmMain.frx":1EA79
      SubMenuID_55_0  =   0
      MenuID_56       =   55
      MenuText_56     =   "�������� (&S)       Ctrl+Shift+F5"
      MenuVisible_56  =   -1  'True
      MenuIcon_56     =   "frmMain.frx":1EB51
      SubMenuID_56_0  =   0
      MenuID_57       =   56
      MenuText_57     =   "-"
      MenuVisible_57  =   -1  'True
      MenuIcon_57     =   "frmMain.frx":1EB71
      SubMenuID_57_0  =   0
      MenuID_58       =   57
      MenuText_58     =   "�����ִ��         F11"
      MenuVisible_58  =   -1  'True
      MenuIcon_58     =   "frmMain.frx":1EB91
      SubMenuID_58_0  =   0
      MenuID_59       =   58
      MenuText_59     =   "�����ִ��         F10"
      MenuVisible_59  =   -1  'True
      MenuIcon_59     =   "frmMain.frx":1EBB1
      SubMenuID_59_0  =   0
      MenuID_60       =   59
      MenuText_60     =   "ִ�е�����         Shift+F11"
      MenuVisible_60  =   -1  'True
      MenuIcon_60     =   "frmMain.frx":1EBD1
      SubMenuID_60_0  =   0
      MenuID_61       =   60
      MenuText_61     =   "����"
      MenuVisible_61  =   -1  'True
      MenuIcon_61     =   "frmMain.frx":1EBF1
      SUBMENU_ITEM_COUNT_61=   5
      SubMenuID_61_0  =   0
      SubMenuText_61_1=   "���ڹ��� (&W)"
      SubMenuID_61_1  =   62
      SubMenuText_61_2=   "��Ϣ���� (&M)"
      SubMenuID_61_2  =   63
      SubMenuText_61_3=   "���� (&P)"
      SubMenuID_61_3  =   64
      SubMenuText_61_4=   "-"
      SubMenuID_61_4  =   65
      SubMenuText_61_5=   "���� (&O)"
      SubMenuID_61_5  =   66
      MenuID_62       =   61
      MenuText_62     =   "���ڹ��� (&W)"
      MenuVisible_62  =   -1  'True
      MenuIcon_62     =   "frmMain.frx":1EC11
      SubMenuID_62_0  =   0
      MenuID_63       =   62
      MenuText_63     =   "��Ϣ���� (&M)"
      MenuVisible_63  =   -1  'True
      MenuIcon_63     =   "frmMain.frx":1ECF5
      SubMenuID_63_0  =   0
      MenuID_64       =   63
      MenuText_64     =   "���� (&P)"
      MenuVisible_64  =   -1  'True
      MenuIcon_64     =   "frmMain.frx":1EDC9
      SubMenuID_64_0  =   0
      MenuID_65       =   64
      MenuText_65     =   "-"
      MenuVisible_65  =   -1  'True
      MenuIcon_65     =   "frmMain.frx":1F042
      SubMenuID_65_0  =   0
      MenuID_66       =   65
      MenuText_66     =   "���� (&O)"
      MenuVisible_66  =   -1  'True
      MenuIcon_66     =   "frmMain.frx":1F062
      SubMenuID_66_0  =   0
      MenuID_67       =   66
      MenuText_67     =   "����"
      MenuVisible_67  =   -1  'True
      MenuIcon_67     =   "frmMain.frx":1F263
      SUBMENU_ITEM_COUNT_67=   3
      SubMenuID_67_0  =   0
      SubMenuText_67_1=   "�����ĵ� (&D)       F1"
      SubMenuID_67_1  =   68
      SubMenuText_67_2=   "ʾ������ (&E)"
      SubMenuID_67_2  =   69
      SubMenuText_67_3=   "�����Ͽؼ��� (&A) Ctrl+F1"
      SubMenuID_67_3  =   70
      MenuID_68       =   67
      MenuText_68     =   "�����ĵ� (&D)       F1"
      MenuVisible_68  =   -1  'True
      MenuIcon_68     =   "frmMain.frx":1F283
      SubMenuID_68_0  =   0
      MenuID_69       =   68
      MenuText_69     =   "ʾ������ (&E)"
      MenuVisible_69  =   -1  'True
      MenuIcon_69     =   "frmMain.frx":1F373
      SubMenuID_69_0  =   0
      MenuID_70       =   69
      MenuText_70     =   "�����Ͽؼ��� (&A) Ctrl+F1"
      MenuVisible_70  =   -1  'True
      MenuIcon_70     =   "frmMain.frx":1F54E
      SubMenuID_70_0  =   0
   End
   Begin VB.PictureBox picToolBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00302D2D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   16845
      TabIndex        =   2
      Top             =   804
      Width           =   16845
   End
   Begin VB.PictureBox picClientArea 
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
      Height          =   5040
      Left            =   0
      ScaleHeight     =   5040
      ScaleWidth      =   16845
      TabIndex        =   0
      Top             =   1200
      Width           =   16845
      Begin VB.PictureBox picWindowClientArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   8880
         ScaleHeight     =   2055
         ScaleWidth      =   5655
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   5655
         Begin DragControlsIDE.TabBar TabBar 
            Height          =   3615
            Left            =   600
            TabIndex        =   5
            Top             =   480
            Visible         =   0   'False
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   6376
         End
      End
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorderSizer 
      Left            =   16200
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      Thickness       =   3
      MinWidth        =   400
      MinHeight       =   100
      Transparency    =   1
      UseSetParent    =   0   'False
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   15600
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      MinWidth        =   400
      MinHeight       =   100
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   1
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
      Caption         =   "�Ͽؼ���"
      MaxButtonVisible=   0   'False
      MinButtonVisible=   0   'False
      BindCaption     =   -1  'True
      Picture         =   "frmMain.frx":1FC6C
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   14160
      Top             =   7320
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DockingPane 
      Left            =   14880
      Top             =   7320
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
'����:      ����, Error 404
'�ļ�:      frmMain.frm
'====================================================

Option Explicit

'��ȡ���������С��״̬
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

'��������
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

'��������
'ֵ     ����
'0      δ�������̣�������������
'1      ���ڳ���
'2      ����̨����
'3      �հ�C++����
Public ProjectType          As Integer

'��ǰ����״̬
'ֵ     ����
'0      ���״̬
'1      ������
'2      �ж�
Public CurrState            As Byte

Public WindowObj            As Object                                                   '��������
Dim NewCreateWindow         As frmCreateOptions                                         '���½���Ŀ������

Public GdbPipe              As clsPipe                                                  'gdb���Թܵ�
Attribute GdbPipe.VB_VarHelpID = -1

'����:      ��Ŀ������˳�֮�������β����
'����:      ExitCode: ������˳���
Private Sub ProcessExitedHandler(ExitCode As Long)
    Dim i                   As Long
    
    Me.tmrCheckProcess.Enabled = False                                                  'ֹͣ��ʱ��
    CurrState = 0                                                                       '���µ���״̬
    For i = 0 To UBound(CurrentProject.Files)                                           '�������ļ����ж������
        CurrentProject.Files(i).TargetWindow.BreakLine = -1
        Call CurrentProject.Files(i).TargetWindow.RedrawBreakpoints
    Next i
    
    frmOutput.OutputLog Lang_Main_Debug_Returned & ExitCode & "(0x" & Hex(ExitCode) & ")"
                
    GdbPipe.DosInput "q" & vbCrLf                                                       '�رչܵ�
    Call ClearDebugWindows(True)                                                        '������е��Դ��ڵ���Ϣ
    
    '�����˵���
    Call AdjustRuntimeMenu
End Sub

'����:      �ô��봰����ʾָ�����ļ�
'����:      FileIndex: ��ѡ�ġ�ָ���ļ���CurrentProject.Files������; �����ָ�������������Ҫָ��FileName����
'.          FileName: ��ѡ�ġ��ļ���; �����ָ�������������Ҫָ��FileIndex����
'����ֵ:    ����ļ��ɹ���ʾ���򷵻ض�Ӧ�Ĵ��봰��; ���򷵻�Nothing
'ע��:      ��������ֻ��ָ������һ�������ͬʱ��ָ��������FileIndex��ִ��
Public Function ShowCodeWindow(Optional FileIndex As Long = -1, Optional FileName As String = "") As frmCodeWindow
    On Error Resume Next
    Dim NewCodeWindow   As frmCodeWindow
    Dim FileData        As String
    Dim tmpData         As String
    
    Set ShowCodeWindow = Nothing
    
    '����Ƿ��ṩ��FileIndex����
    If FileIndex <> -1 Then
        With CurrentProject.Files(FileIndex)
            '���û�ж�Ӧ�Ĵ��봰�ھʹ���һ���µģ��еĻ����л���ȥ
            If .TargetWindow Is Nothing Then
                Set NewCodeWindow = CreateNewCodeWindow(FileIndex)                  '�����µĴ��봰�岢���ð󶨵��ļ����
                NewCodeWindow.Caption = GetFileName(.FilePath)
                
                Err.Clear
                Open .FilePath For Input As #1                                      '���Դ򿪶�Ӧ�Ĵ����ļ�
                    If Err.Number <> 0 Then
                        Close #1
                        Exit Function
                    Else
                        Do While Not EOF(1)
                            Line Input #1, tmpData
                            FileData = FileData & tmpData & vbCrLf
                        Loop
                    End If
                Close #1
                
                NewCodeWindow.SyntaxEdit.Text = FileData
                Me.TabBar.AddForm NewCodeWindow
            Else
                Me.TabBar.SwitchToByForm .TargetWindow
            End If
            Set ShowCodeWindow = .TargetWindow
        End With
        Exit Function
    End If
    
    '����Ƿ��ṩ��FileName����
    If FileName <> "" Then
        Dim i   As Long
        
        '���Ҹ��ļ���Ӧ�Ĵ����ļ�
        For i = 0 To UBound(CurrentProject.Files)
            With CurrentProject.Files(i)
                If .FilePath = FileName Then
                    If .TargetWindow Is Nothing Then                                    '�ô����ļ�û���Ѵ򿪵Ĵ��봰��
                        Set NewCodeWindow = CreateNewCodeWindow(i)                          '�����µĴ��봰�岢���ð󶨵��ļ����
                        NewCodeWindow.Caption = GetFileName(.FilePath)
                        
                        Err.Clear
                        Open .FilePath For Input As #1                                      '���Դ򿪶�Ӧ�Ĵ����ļ�
                            If Err.Number <> 0 Then
                                Close #1
                                Exit Function
                            Else
                                Do While Not EOF(1)
                                    Line Input #1, tmpData
                                    FileData = FileData & tmpData & vbCrLf
                                Loop
                            End If
                        Close #1
                        
                        NewCodeWindow.SyntaxEdit.Text = FileData
                        Me.TabBar.AddForm NewCodeWindow
                    Else
                        Me.TabBar.SwitchToByForm .TargetWindow
                    End If
                    Set ShowCodeWindow = .TargetWindow
                    Exit Function
                End If
            End With
        Next i
    End If
End Function

'����:      ���ݲ�ͬ������״̬�����˵�״̬
Private Sub AdjustRuntimeMenu()
    '0: ���ģʽ; 1: ������; 2: �ж�
    Select Case CurrState
        Case 0                                                          '���ģʽ
            Me.DarkMenu.MenuEnabled(52) = True                              '����
            Me.DarkMenu.MenuText(52) = Lang_Main_Run_Menu_Start
            Me.DarkMenu.MenuEnabled(53) = False                             '�ж�
            Me.DarkMenu.MenuEnabled(54) = False                             'ֹͣ
            Me.DarkMenu.MenuEnabled(55) = False                             '��������
            Me.DarkMenu.MenuEnabled(57) = False                             '�����
            Me.DarkMenu.MenuEnabled(58) = False                             '�����
            Me.DarkMenu.MenuEnabled(59) = False                             'ִ�е�����
        
        Case 1                                                          '������
            Me.DarkMenu.MenuEnabled(52) = False
            Me.DarkMenu.MenuText(52) = Lang_Main_Run_Menu_Continue
            Me.DarkMenu.MenuEnabled(53) = True
            Me.DarkMenu.MenuEnabled(54) = True
            Me.DarkMenu.MenuEnabled(55) = True
            Me.DarkMenu.MenuEnabled(57) = False
            Me.DarkMenu.MenuEnabled(58) = False
            Me.DarkMenu.MenuEnabled(59) = False
        
        Case 2                                                          '�ж�
            Me.DarkMenu.MenuEnabled(52) = True
            Me.DarkMenu.MenuText(52) = Lang_Main_Run_Menu_Continue
            Me.DarkMenu.MenuEnabled(53) = False
            Me.DarkMenu.MenuEnabled(54) = True
            Me.DarkMenu.MenuEnabled(55) = True
            Me.DarkMenu.MenuEnabled(57) = True
            Me.DarkMenu.MenuEnabled(58) = True
            Me.DarkMenu.MenuEnabled(59) = True
        
    End Select
End Sub

'����:      ������е��Դ����������Ϣ
'����:      ClearBreakpoints: ��ѡ�ġ�ָ���Ƿ���նϵ��б�����ĵ�ַ��ͨ���ڵ����ڼ䲻��Ҫ��նϵ��б��ڵ�����ɺ����Ҫ
Private Sub ClearDebugWindows(Optional ClearBreakpoints As Boolean = False)
    If ClearBreakpoints Then                                                        '�ϵ�
        Call frmBreakpoints.ClearEverything
    End If
    Call frmLocals.ClearEverything                                                  '����
    Call frmCallStack.ClearEverything                                               '���ö�ջ
End Sub

'����:      ��鵱ǰ�Ƿ���δ������ļ�
'����ֵ:    �����δ������ļ����򷵻�True
Private Function IsSaveRequired() As Boolean
    'On Error Resume Next       'todo
    
    IsSaveRequired = False
    If CurrentProject.Changed Then                                              '�����ļ���Ҫ����
        IsSaveRequired = True
    Else                                                                        '������һ�������ļ���Ҫ����
        Dim i               As Long
        
        For i = 0 To UBound(CurrentProject.Files)
            If CurrentProject.Files(i).Changed Then
                IsSaveRequired = True
                Exit For
            End If
        Next i
    End If
End Function

'����:      ��������Ŀ���˵�
Private Sub mnuOpen_Click()
    'ToDo
    NoSkinMsgBox ShowOpen(Me.hwnd, "Dilidi - Open", "ϴƨƨ�ļ�(*.cpp)" & vbNullChar & "*.cpp")
End Sub

'����:      �����桱�˵�
'����ֵ:    1=����ɹ�; 2=����ʧ��; 3=ȡ��; 4=������
Private Function mnuSave_Click() As Integer
    Dim i                   As Long
    
    frmSaveBox.InitFileIndexMap                                                         '��ʼ�����ӳ���
    If CurrentProject.Changed Then                                                      '��ǰ�����ļ�������
        frmSaveBox.AddFileIndexMap -1
    End If
    For i = 0 To UBound(CurrentProject.Files)                                           '��黹û�б�����ļ�
        If CurrentProject.Files(i).Changed Then                                             '��鵽��û�б�����ļ�
            frmSaveBox.AddFileIndexMap i
        End If
    Next i
    If frmSaveBox.lstFiles.ListCount = 0 Then                                           '���û���ļ���Ҫ����
        Exit Function
    End If
    For i = 0 To frmSaveBox.lstFiles.ListCount - 1                                      '��ѡ�����ļ�
        frmSaveBox.lstFiles.Selected(i) = True
    Next i
    frmSaveBox.bSaveFlag = 0                                                            '��ʼ��������
    frmSaveBox.bBlock = True                                                            '��������ִ��
    Me.Enabled = False
    frmSaveBox.Show
    SetWindowPos frmSaveBox.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE    '�ñ���Ի����ö�
    Do                                                                                  '�ȴ�����Ի���ر�
        Sleep 50
        DoEvents
    Loop Until Not frmSaveBox.bBlock
    mnuSave_Click = frmSaveBox.bSaveFlag                                                '���ر�����
    If mnuSave_Click = 0 Then                                                           'û��ѡ���������Ϊȡ��
        mnuSave_Click = 3
    End If
    Unload frmSaveBox                                                                   'ж�ص�����Ի����ͷ���Դ
End Function

'����:      �����Ϊ���˵�
Private Sub mnuSaveAs_Click()
    NoSkinMsgBox ShowSave(Me.hwnd, "Shar.cpp", "Save", "fsaf(*.cpp)" & vbNullChar & "*.cpp")
End Sub

'����:      ���½���Ŀ���˵�
Private Sub mnuNewProject_Click()
    If Not NewCreateWindow Is Nothing Then                                              'ж�ص���һ�����½���Ŀ������
        Unload NewCreateWindow
        Set NewCreateWindow = Nothing
    End If
    Set NewCreateWindow = New frmCreateOptions
    Unload frmCreate
    Me.Enabled = False
    Me.DarkWindowBorderSizer.Bind = False
    SetParent NewCreateWindow.hwnd, 0
    NewCreateWindow.Move Screen.Width / 2 - frmCreate.Width / 2, Screen.Height / 2 - frmCreate.Height / 2
    NewCreateWindow.DarkTitleBar_NoDrop.Visible = True
    NewCreateWindow.DarkWindowBorder.Bind = True
    NewCreateWindow.Show
    NewCreateWindow.TypeOption(1).Focused = True
End Sub

'����:      �����С��˵�
Private Sub mnuRun_Click()
    On Error Resume Next
    
    Dim GccPipe             As New clsPipe                                      'g++�ܵ�
    Dim GccCmdLine          As String                                           'g++������
    Dim ExePath             As String                                           'exe�ļ�����·��
    Dim PipeOutput          As String                                           '�ܵ����������
    Dim GccOutputContent()  As String                                           '���зֿ���g++�������
    Dim i                   As Long
    Dim MsgBoxRtn           As VbMsgBoxResult                                   '����ȷ�Ͽ�ķ���ֵ
    Dim SaveRtn             As Integer                                          '���淵��ֵ
    
    Call AdjustRuntimeMenu
    Me.DarkMenu.MenuEnabled(52) = False                                         '�������в˵�
    
    If CurrState = 2 Then                                                       '�����ж�״̬
        Call ClearDebugWindows                                                      '������е��Դ��ڵ���Ϣ
        GdbPipe.DosInput "continue" & vbCrLf                                        '���ͼ�����������
        CurrState = 1                                                               '���µ���״̬
        Call AdjustRuntimeMenu                                                      '�����˵�
        Exit Sub
    End If
    
    If IsSaveRequired() Then                                                    '��ʾ�����ļ�
        MsgBoxRtn = NoSkinMsgBox(Lang_Main_SaveBeforeCompile, vbQuestion Or vbYesNoCancel, Lang_Msgbox_Confirm)
        If MsgBoxRtn = vbYes Then
            SaveRtn = mnuSave_Click()
            If SaveRtn = 2 Then                                                         '����ʱ����
                If NoSkinMsgBox(Lang_Main_SaveFailedBeforeCompile, vbQuestion Or vbYesNo, Lang_Msgbox_Confirm) = vbNo Then
                    frmOutput.OutputLog Lang_Main_DebugAborted
                    Call AdjustRuntimeMenu
                    Exit Sub
                End If
            ElseIf SaveRtn = 3 Or SaveRtn = 4 Then                                      '�û�ȡ������ ���� �û�ѡ�񲻱��� ��ȡ���������Ĳ���
                frmOutput.OutputLog Lang_Main_DebugAborted
                Call AdjustRuntimeMenu
                Exit Sub
            End If
        ElseIf MsgBoxRtn = vbCancel Then
            frmOutput.OutputLog Lang_Main_DebugAborted                              '�û�ѡ��ȡ������
            Call AdjustRuntimeMenu
            Exit Sub
        End If
    End If
    '======================================================================
    
    ExePath = ProjectFolderPath & CurrentProject.ProjectName & ".exe"
    If Dir(ExePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem) <> "" Then  '��⵽ͬ����exe�ļ�
        If NoSkinMsgBox(Lang_Main_ReplaceExe_1 & ExePath & Lang_Main_ReplaceExe_2, vbQuestion Or vbYesNo, Lang_Msgbox_Confirm) = vbYes Then
            Kill ExePath                                                                'ɾ����ͬ���ļ�
        Else
            frmOutput.OutputLog Lang_Main_DebugAborted
            Call AdjustRuntimeMenu
            Exit Sub
        End If
    End If
    Call frmOutput.ClearEverything                                              '������
    Call frmErrorList.ClearEverything                                           '��մ����б�
    '======================================================================
    
    'ʹ��g++���б���
    '                   ��ת����ǰ�������ڵ��̷�                    ������g++.exe���б���      ������Ϊ���Գ���           �����е�cpp�����ļ�
    '�����ʽ: cmd /c ���̷���: && cd "��g++.exe����Ŀ¼��" && "��g++.exe·����" [-mwindows] -g -o "�����·����" "��cpp�ļ�1��" "��cpp�ļ�2��"
    '                                       ��ת��g++.exe���ڵ�Ŀ¼                 ���Ƿ�Ϊ�����г���   �������EXE���·��
    GccCmdLine = "cmd /c " & Left(GccPath, 1) & ": && " & _
        "cd """ & Left(GccPath, InStrRev(GccPath, "\")) & """ && " & _
        """" & GccPath & """ -static -g -o """ & ExePath & """"
    For i = 0 To UBound(CurrentProject.Files)
        If Not Right(CurrentProject.Files(i).FilePath, 2) = ".h" And Not Right(CurrentProject.Files(i).FilePath, 4) = ".hpp" Then
            GccCmdLine = GccCmdLine & " """ & CurrentProject.Files(i).FilePath & """"
        End If
    Next i
    If GccPipe.InitDosIO(GccCmdLine) = 0 Then                                   'g++�ܵ�����ʧ��
        frmOutput.OutputLog Lang_Main_GccStartFailed & GccCmdLine
        Call AdjustRuntimeMenu
        Exit Sub
    End If
    frmOutput.OutputLog Lang_Main_StartingGcc & GccCmdLine
    Do While ProcessExists(GccPipe.hProcess)                                    '�ȴ�g++ִ�����
        Sleep 50
        DoEvents
    Loop
    GccPipe.DosOutput PipeOutput, vbNullChar & vbNullChar                       '��ȡg++���
    GccOutputContent = Split(PipeOutput, vbCrLf)
    If UBound(GccOutputContent) >= 0 Then
        For i = 0 To UBound(GccOutputContent)                                   '�������
            If GccOutputContent(i) <> "" Then                                       '������ǿ���
                frmOutput.OutputLog GccOutputContent(i)
                If GccOutputContent(i) Like "*:\*:*:*: *" And InStr(Left(GccOutputContent(i), 5), ":\") <> 0 Then   '��x:\File Name:Line:Column: error reason��
                    Dim StrPos          As Long                                             '�����ַ����Ľ��
                    Dim CurrFileName    As String                                           '��Ӧ���ļ�
                    Dim CurrLineNumber  As Long                                             '��Ӧ���к�
                    Dim CurrColNumber   As Long                                             '��Ӧ���к�
                    Dim InfoType        As Byte                                             '����Ϣ�����ͣ�0: error; 1: warning; 2: info��
                    
                    '����g++�������¼��frmOutput������ж�Ӧ��Ϣ��
                    StrPos = InStr(GccOutputContent(i), ":\")
                    StrPos = InStr(StrPos + 2, GccOutputContent(i), ":")
                    CurrFileName = Left(GccOutputContent(i), StrPos - 1)                    '��[x:\File Name]:Line:Column: error reason��
                    GccOutputContent(i) = Right(GccOutputContent(i), _
                        Len(GccOutputContent(i)) - Len(CurrFileName) - 1)                   '��x:\File Name:[Line:Column: error reason]��
                    StrPos = InStr(GccOutputContent(i), ":")
                    CurrLineNumber = CLng(Left(GccOutputContent(i), StrPos - 1))            '��[Line]:Column: error reason��
                    GccOutputContent(i) = Right(GccOutputContent(i), _
                        Len(GccOutputContent(i)) - Len(CStr(CurrLineNumber)) - 1)           '��Line:[Column: error reason]��
                    StrPos = InStr(GccOutputContent(i), ":")
                    CurrColNumber = CLng(Left(GccOutputContent(i), StrPos - 1))             '��[Column]: error reason��
                    GccOutputContent(i) = Right(GccOutputContent(i), _
                        Len(GccOutputContent(i)) - Len(CStr(CurrColNumber)) - 2)            '��Column: [error reason]��
                    frmOutput.AddLineInfo False, CurrFileName, CurrLineNumber, CurrColNumber
                    
                    '��ӵ������б���
                    GccOutputContent(i) = Trim(GccOutputContent(i))                         'ȥ����������ǰ��Ŀո�
                    InfoType = 0                                                            'Ĭ������ϢΪ����
                    If LCase(Left(GccOutputContent(i), 7)) = "warning" Then                 '��⵽��ϢΪwarning����
                        InfoType = 1
                    ElseIf LCase(Left(GccOutputContent(i), 4)) = "note" Then                '��⵽��ϢΪnote����
                        InfoType = 2
                    Else
                        If LCase(Left(GccOutputContent(i), 5)) <> "error" Then
                            Stop
                        End If
                    End If
                    frmErrorList.AddErrorInfoListItem InfoType, GccOutputContent(i), CurrFileName, CurrLineNumber, CurrColNumber
                End If
            End If
            DoEvents                                                                    '��Ҫ�����߳�
        Next i
        Call frmErrorList.AddErrorListItem                                          '��Ӵ�����Ϣ���б���
    End If
    If Dir(ExePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem) = "" Then   '���exe·�������ڣ���˵�����벻�ɹ�
        frmOutput.OutputLog Lang_Main_CompileFailed
        Call AdjustRuntimeMenu
        Exit Sub
    Else
        frmOutput.OutputLog Lang_Main_CompileSucceed & ExePath
        frmOutput.AddLineInfo True, ExePath, 0
    End If
    '======================================================================
    
    '���������Խ��̡��ý�������֮�����𣬵ȴ�gdb����
    Dim si                  As STARTUPINFO                                      '����������Ϣ
    Dim sa                  As SECURITY_ATTRIBUTES                              '��ȫ����
    
    With sa                                                                     '���ð�ȫ����
        .bInheritHandle = 1
        .lpSecurityDescriptor = 0
        .nLength = Len(sa)
    End With
    If CreateProcess(0, ExePath, sa, sa, ByVal 1, _
        NORMAL_PRIORITY_CLASS Or CREATE_SUSPENDED, ByVal 0, ByVal 0, si, DebugProgramInfo) <> 1 Then
        
        frmOutput.OutputLog Lang_Main_RunFailed & ExePath & " (" & Err.LastDllError & ")"
        frmOutput.AddLineInfo True, ExePath, 0
        Call AdjustRuntimeMenu
        Exit Sub
    End If
    frmOutput.OutputLog Lang_Main_RunSucceed & DebugProgramInfo.dwProcessId & "(" & Hex(DebugProgramInfo.dwProcessId) & ")"
    '======================================================================
    
    '����gdb�ܵ�
    Set GdbPipe = New clsPipe
    If GdbPipe.InitDosIO(GdbPath & " -q -nw") = 0 Then                          '����gdb���Թܵ�ʧ��
        TerminateProcess DebugProgramInfo.hProcess, 0                               'ɱ�������Խ��̣���������
        Set GdbPipe = Nothing                                                       '�ر�gdb�ܵ�
        frmOutput.OutputLog Lang_Main_GdbFailed
        Call AdjustRuntimeMenu
        Exit Sub
    End If
    frmOutput.OutputLog Lang_Main_GdbSucceed & GdbPipe.dwProcessId & "(" & Hex(GdbPipe.dwProcessId) & ")"
    '======================================================================
    
    frmOutput.OutputLog Lang_Main_GdbLoadingSymbols_1 & ExePath & Lang_Main_GdbLoadingSymbols_2
    frmOutput.AddLineInfo True, ExePath, 0
    GdbPipe.DosInput "file """ & Replace(ExePath, "\", "/") & """" & vbCrLf     '��exe�ļ���ȡ����
    GdbPipe.DosOutput PipeOutput, "(gdb) ", 5000                                '��ȡgdb�����
    If InStr(PipeOutput, "no debugging symbols found") <> 0 Or _
        InStr(PipeOutput, "No such file or directory") <> 0 Then                    'gdb�����no debugging symbols found�����ߡ�No such file or directory�������ط���ʧ��
        frmOutput.OutputLog CStr(Split(PipeOutput, vbCrLf)(0))                      '������ط��ŵĴ���
        If NoSkinMsgBox(Lang_Main_GdbLoadSymbolsFailure_1 & ExePath & Lang_Main_GdbLoadSymbolsFailure_2, vbExclamation Or vbYesNo, Lang_Msgbox_Confirm) = vbNo Then
            TerminateProcess DebugProgramInfo.hProcess, 0                               'ɱ�������Խ��̣���������
            Set GdbPipe = Nothing                                                       '�ر�gdb�ܵ�
            frmOutput.OutputLog Lang_Main_DebugAborted
            Call AdjustRuntimeMenu
            Exit Sub
        End If
    End If
    '======================================================================
    
    GdbPipe.DosInput "set pagination off" & vbCrLf                              '�ر�gdb��"Type to continue, or q to quit"��Ϣ
    GdbPipe.DosInput "set print repeats 0" & vbCrLf                             '�ر�gdb�����ظ�������Ԫ�صġ�<repeats n times>�����
    '======================================================================
    
    frmOutput.OutputLog Lang_Main_GdbAttaching
    GdbPipe.DosInput "attach " & DebugProgramInfo.dwProcessId & vbCrLf          '���ӵ������Խ���
    GdbPipe.DosOutput PipeOutput, "(gdb) ", 5000                                '��ȡgdb�����
    If InStr(PipeOutput, "Can't attach") <> 0 Then                              'gdb�����Can't attach to process.�������ӽ���ʧ��
        TerminateProcess DebugProgramInfo.hProcess, 0                               'ɱ�������Խ��̣���������
        Set GdbPipe = Nothing                                                       '�ر�gdb�ܵ�
        frmOutput.OutputLog Lang_Main_GdbAttachFailed_1 & DebugProgramInfo.dwProcessId & "(" & Hex(DebugProgramInfo.dwProcessId) & ") " & Lang_Main_GdbAttachFailed_2
        frmOutput.OutputLog Lang_Main_DebugAborted
        Call AdjustRuntimeMenu
        Exit Sub
    End If
    GdbPipe.DosOutput PipeOutput, "(gdb) ", 5000                                '�ȴ�gdb�����ɣ���ʱ5��
    '======================================================================
    
    Dim j                   As Long
    Dim k                   As Long
    Dim CurrFilePath        As String                                           '�����С�\���滻�ɡ�/�����ļ�·�������ԡ�:����β���Ա�����к�
    Dim SplitTmp()          As String                                           '�ַ����ָ��
    Dim BreakpointAdded     As Boolean                                          '�ϵ��Ƿ�ɹ����
    
    ReDim GdbBreakpoints(0)                                                     '��ʼ���ϵ�ӳ���
    For i = 0 To UBound(CurrentProject.Files)                                   '��������ļ��Ķϵ�
        CurrFilePath = Replace(CurrentProject.Files(i).FilePath, "\", "/") & """:"
        For j = 0 To UBound(CurrentProject.Files(i).Breakpoints) - 1                'Ϊɶ - 1: ��Ϊ�ϵ��б�������һ����û�õ�
            GdbPipe.ClearPipe                                                           '����ܵ�
            GdbPipe.DosInput "b """ & CurrFilePath & CStr(CurrentProject.Files(i).Breakpoints(j).CodeLn) & vbCrLf
            GdbPipe.DosOutput PipeOutput, "(gdb) "                                      '��ȡgdb�����
            SplitTmp = Split(PipeOutput, vbCrLf)                                        '��������зֿ�
            BreakpointAdded = False                                                     '�ȰѶϵ�ɹ���ӱ��ΪFalse
            For k = 0 To UBound(SplitTmp)                                               '���з���
                PipeOutput = SplitTmp(k)
                If PipeOutput Like "Breakpoint * at *, line *" Then                         '��Ӷϵ�ɹ����ȡ�ϵ���Ϣ
                    BreakpointAdded = True
                    SplitTmp = Split(Split(PipeOutput, "Breakpoint ")(1), " at ")
                    If CLng(SplitTmp(0) > UBound(GdbBreakpoints)) Then                          '���gdb�ϵ�ӳ���Ĵ�С�Ƿ��㹻������������һ��
                        ReDim Preserve GdbBreakpoints(CLng(SplitTmp(0)))
                    End If
                    GdbBreakpoints(CLng(SplitTmp(0))).FileIndex = i                             '��¼gdb�ϵ�����Ӧ���ļ���źͶϵ����
                    GdbBreakpoints(CLng(SplitTmp(0))).BreakpointIndex = j
                    frmBreakpoints.lvBreakpoints.SetItemText CStr(Split(SplitTmp(1), ": file")(0)), CurrentProject.Files(i).Breakpoints(j).ListViewIndex, 2
                    Exit For
                ElseIf PipeOutput Like "No line * in file *" Then                           'û��ָ�����кţ���No line * in file "*".����
                    Dim tmpFileLine As Long
                    
                    tmpFileLine = Replace(Split(PipeOutput, " in file """)(0), "No line ", "")  '����No line [*] in file "*".����
                    frmOutput.OutputLog Lang_Main_GdbBreakpointError_1 & CurrentProject.Files(i).FilePath & _
                        Lang_Main_GdbBreakpointError_2 & tmpFileLine & Lang_Main_GdbBreakpointError_3
                    frmOutput.AddLineInfo False, CurrentProject.Files(i).FilePath, tmpFileLine
                End If
            Next k
            If Not BreakpointAdded Then
                frmBreakpoints.lvBreakpoints.SetItemText Lang_Main_GdbBreakpoint_Invalid, CurrentProject.Files(i).Breakpoints(j).ListViewIndex, 2
            End If
        Next j
    Next i
    '======================================================================
    
    GdbPipe.DosInput "continue" & vbCrLf                                        'ʹĿ����̼�������
    ResumeThread DebugProgramInfo.hThread                                       '����ִ��Ŀ����̵����߳�
    frmOutput.OutputLog Lang_Main_RunningInfo_1 & DebugProgramInfo.dwProcessId & "(" & Hex(DebugProgramInfo.dwProcessId) & ") " & Lang_Main_RunningInfo_2
    CurrState = 1                                                               '���µ���״̬
    Call AdjustRuntimeMenu
    Me.tmrCheckProcess.Enabled = True                                           '��ʼ�ȴ����̽���
End Sub

'����:      ������������
Public Sub HideStartupPage()
    On Error Resume Next
    Unload NewCreateWindow
    Me.TabBar.Visible = True
    
    Me.DarkMenu.MenuEnabled(3) = True                                                   '����
    Me.DarkMenu.MenuEnabled(4) = True                                                   '���Ϊ
    Me.DarkMenu.MenuEnabled(7) = True                                                   '�༭
    Me.DarkMenu.MenuEnabled(27) = True                                                  '��ͼ
    Me.DarkMenu.MenuEnabled(34) = True                                                  '����
    Me.DarkMenu.MenuEnabled(37) = True                                                  '����
End Sub

'����:      ��ʾ��������
Public Sub ShowStartupPage()
    frmCreate.DarkTitleBar_NoDrop.Visible = False                                       '����ʾ�������ͱ߿�
    frmCreate.DarkWindowBorder.Bind = False
    SetParent frmCreate.hwnd, Me.picClientArea.hwnd                                     '�á��½���Ŀ����Ϊ��������Ӵ���
    frmCreate.Move 0, 0                                                                 '������λ��
    frmCreate.Show
End Sub

Private Sub DarkMenu_MenuItemClicked(MenuID As Integer)
    Me.DarkMenu.HideMenu                                                            '���²˵�����������ز˵�
    Select Case MenuID
        Case 1                                                                          '�½�
            Call mnuNewProject_Click
        
        Case 2                                                                          '����
            Call mnuOpen_Click
        
        Case 3                                                                          '����
            Call mnuSave_Click
        
        Case 4                                                                          '���Ϊ
            Call mnuSaveAs_Click
        
        Case 6                                                                          '�˳�
            Unload Me
        
        Case 32                                                                         '�����б�
            Me.DockingPane.ShowPane 4
        
        Case 39                                                                         '�ϵ��б�
            Me.DockingPane.ShowPane 6
        
        Case 42                                                                         '����
            Me.DockingPane.ShowPane 8
        
        Case 45                                                                         '���ö�ջ
            Me.DockingPane.ShowPane 10
        
        Case 52                                                                         '����
            Call mnuRun_Click
        
    End Select
End Sub

Private Sub DockingPane_Resize()
    'On Error Resume Next       'todo
    
    If ProjectType <> 0 Then                                                            '�����������������Ļ��͵������ڻ�ͻ�����С
        Dim cLeft   As Long, cRight   As Long, cTop   As Long, cBottom   As Long
        
        Me.DockingPane.GetClientRect cLeft, cTop, cRight, cBottom
        Me.picWindowClientArea.Move cLeft, cTop, cRight - cLeft, cBottom - cTop
        Me.TabBar.Move 0, 0, Me.picWindowClientArea.ScaleWidth, Me.picWindowClientArea.ScaleHeight
        
        Call Form_Resize                                                                    '������ڿͻ�����������󻯵Ĵ��ڣ������С���е���
    End If
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    
    '����Logo
    frmStartupLogo.Show
    SetWindowPos frmStartupLogo.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    frmStartupLogo.SetFocus
    frmStartupLogo.Refresh
    
    '����������ʾ�ı�����
    If CreateToolTip() = 0 Then
        MsgBox "����������ʾ�ı�����ʧ�ܣ�", vbCritical, "����"
    End If

    '�����ַ�����Դ������û��ؼ�������أ��û��ؼ�����ʹ����Щ�ַ�����Դ�����Ƿ���Initialize�¼�������Load�¼�
    '�����ַ�����Դ
    If Not LoadLanguage(1001) Then
        MsgBox "�����ַ�����Դʧ�ܣ�" & Err.Number & ": " & Err.Description, vbCritical, "����"
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    '�������ش��壬��һ�м���������ʾ
    Me.Hide
    Me.DarkTitleBar.MinButtonVisible = True
    Me.DarkTitleBar.MaxButtonVisible = True
    Me.Caption = Lang_Application_Title
    
    '���ز˵��ı�
    If Not LoadLanguage(1001, True) Then
        MsgBox "�����ַ�����Դʧ�ܣ�" & Err.Number & ": " & Err.Description, vbCritical, Lang_Msgbox_Error
    End If
    
    '�������ͻ�����
    Me.picClientArea.Height = Me.ScaleHeight - Me.picClientArea.Top                                                     '�����ͻ������Ĵ�С
    Me.picWindowClientArea.BackColor = Me.BackColor                                                                     '���ڿͻ�������ɫ
    
    '����������
    Dim ClientHeight        As Integer, ClientWidth             As Integer
    Dim i                   As Integer
    
    Me.DockingPane.AttachToWindow Me.picClientArea.hwnd                                                                 '�󶨹�����
    ClientHeight = Me.picClientArea.ScaleHeight / Screen.TwipsPerPixelY
    ClientWidth = Me.picClientArea.ScaleWidth / Screen.TwipsPerPixelX
    Me.DockingPane.CreatePane 1, 100, ClientHeight, DockLeftOf                                                          '�ؼ���
    Me.DockingPane(1).Handle = frmControlBox.hwnd
    Me.DockingPane(1).Title = Lang_ControlBox_Caption
    Me.DockingPane.CreatePane 2, 250, ClientHeight / 2, DockRightOf                                                     '����
    Me.DockingPane(2).Handle = frmProperties.hwnd
    Me.DockingPane(2).Title = Lang_Properties_Caption
    Me.DockingPane.CreatePane 3, 250, ClientHeight / 2, DockRightOf                                                     '������Դ������
    Me.DockingPane(3).Handle = frmSolutionExplorer.hwnd
    Me.DockingPane(3).Title = Lang_SolutionExplorer_Caption
    Me.DockingPane.CreatePane 4, ClientWidth / 2, 120, DockBottomOf Or DockLeftOf                                       '�����б�
    Me.DockingPane(4).Handle = frmErrorList.hwnd
    Me.DockingPane(4).Title = Lang_ErrorList_Caption
    Me.DockingPane.CreatePane 5, ClientWidth / 2, 120, DockBottomOf Or DockRightOf                                      '���
    Me.DockingPane(5).Handle = frmOutput.hwnd
    Me.DockingPane(5).Title = Lang_Output_Caption
    Me.DockingPane.CreatePane 6, ClientWidth / 2, 120, DockBottomOf Or DockRightOf                                      '�ϵ��б�
    Me.DockingPane(6).Handle = frmBreakpoints.hwnd
    Me.DockingPane(6).Title = Lang_Breakpoints_Caption
    Me.DockingPane.CreatePane 7, ClientWidth / 2, 120, DockBottomOf                                                     '���Ӵ���
    Me.DockingPane(7).Handle = frmWatch.hwnd
    Me.DockingPane(7).Title = Lang_Watch_Caption
    Me.DockingPane.CreatePane 8, ClientWidth / 2, 120, DockBottomOf                                                     '����
    Me.DockingPane(8).Handle = frmLocals.hwnd
    Me.DockingPane(8).Title = Lang_Locals_Caption
    Me.DockingPane.CreatePane 9, ClientWidth / 2, 120, DockBottomOf                                                     '��������
    Me.DockingPane(9).Handle = frmImmediate.hwnd
    Me.DockingPane(9).Title = Lang_Immediate_Caption
    Me.DockingPane.CreatePane 10, ClientWidth / 2, 120, DockBottomOf                                                    '���ö�ջ
    Me.DockingPane(10).Handle = frmCallStack.hwnd
    Me.DockingPane(10).Title = Lang_CallStack_Caption
    Me.DockingPane.CreatePane 11, ClientWidth / 2, 120, DockBottomOf                                                    '�߳�
    Me.DockingPane(11).Handle = frmThreads.hwnd
    Me.DockingPane(11).Title = Lang_Threads_Caption
    Me.DockingPane.CreatePane 12, ClientWidth / 2, 120, DockBottomOf                                                    'ģ��
    Me.DockingPane(12).Handle = frmModules.hwnd
    Me.DockingPane(12).Title = Lang_Modules_Caption
    Me.DockingPane.CreatePane 13, ClientWidth / 2, 250, DockBottomOf                                                    '�ڴ�
    Me.DockingPane(13).Handle = frmMemory.hwnd
    Me.DockingPane(13).Title = Lang_Memory_Caption
    Me.DockingPane.CreatePane 14, ClientWidth / 2, 250, DockBottomOf                                                    '�Ĵ���
    Me.DockingPane(14).Handle = frmRegisters.hwnd
    Me.DockingPane(14).Title = Lang_Registers_Caption
    Me.DockingPane.CreatePane 15, ClientWidth / 2, 250, DockBottomOf                                                    '�����
    Me.DockingPane(15).Handle = frmDisassembly.hwnd
    Me.DockingPane(15).Title = Lang_Disassembly_Caption
    For i = 1 To 15                                                                                                     '�������е�Pane
        Me.DockingPane(i).Close
    Next i
    
    '����Docking Pane����ʽ
    Me.DockingPane.Options.ShowDockingContextStickers = True                                                            '��ʾDocking stickers
    Me.DockingPane.Options.AlphaDockingContext = True                                                                   '�ƶ���ʱ��͸��
    Me.DockingPane.Options.ThemedFloatingFrames = True                                                                  '��Ϊ����ʱ�߿򱣳���ʽ
    Me.DockingPane.Options.ShowContentsWhileDragging = True
    If DockingPaneGlobalSettings.ResourceImages.LoadFromFile(GetAppPath & "Skin.dll", "Office2010Black.ini") = False Then
        MsgBox "������ʽʧ�ܣ�", vbCritical, Lang_Msgbox_Error
    End If
    Me.DockingPane.VisualTheme = ThemeResource                                                                          '����Ϊ����Դ�ļ���ȡ��ʽ
    Me.DockingPane.PaintManager.SplitterSize = 2                                                                        '���÷ָ�����Ĵ�С
    
    'If Not Me.SkinFramework.LoadSkin("Skin.cjstyles", "NormalBlue.ini") Then                                            '����Ƥ�� [ToDo]
        'MsgBox "����Ƥ��ʧ�ܣ�", vbCritical, Lang_Msgbox_Error todo: multi language
    'End If
    
    'todo ɾ��-----------------
    GccPath = "C:\Program Files (x86)\MinGW\bin\g++.exe"
    GdbPath = "C:\Program Files (x86)\MinGW\bin\gdb.exe"
    '--------------------------
    
    '���ò���Ҫ�Ĳ˵�
    Me.DarkMenu.MenuEnabled(3) = False                                                                                  '����
    Me.DarkMenu.MenuEnabled(4) = False                                                                                  '���Ϊ
    Me.DarkMenu.MenuEnabled(7) = False                                                                                  '�༭
    Me.DarkMenu.MenuEnabled(27) = False                                                                                 '��ͼ
    Me.DarkMenu.MenuEnabled(34) = False                                                                                 '����
    Me.DarkMenu.MenuEnabled(37) = False                                                                                 '����
    
    '���ô������໯������������⼰�����������Ҽ��ر�
    Dim lpObj               As Long                                                                                     'ָ�򴰿���������ָ��
    Set WindowObj = Me
    lpObj = ObjPtr(WindowObj)                                                                                           '��ȡָ�򴰿���������ָ��
    SetPropA Me.hwnd, "WindowObj", lpObj                                                                                '��¼���ڵ������ַ�������໯ж�ش�����
    'SetPropA Me.hWnd, "PrevWndProc", SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf MainWindowMaximizeCloseFixProc)    '[ToDo]
    
    '��ʾ����ҳ��
    Call ShowStartupPage
    picToolBar.Move 0, Me.DarkMenu.Top + Me.DarkMenu.Height
    Me.picClientArea.Move 0, Me.picToolBar.Top + Me.picToolBar.Height
    
    '��ʼ�����в˵�
    CurrState = 0
    Call AdjustRuntimeMenu
    
    'ж��LOGO
    Unload frmStartupLogo
    Me.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    
    '�رղ˵�
    Me.DarkMenu.HideMenu
    
    '��顰�½���Ŀ�������Ƿ�ر�
    If frmCreateOptions.Visible = True Then
        frmCreateOptions.Show
        frmCreateOptions.SetFocus
        Cancel = 1
        Exit Sub
    End If
    If Not NewCreateWindow Is Nothing Then
        If NewCreateWindow.Visible = True Then
            NewCreateWindow.Show
            NewCreateWindow.SetFocus
            Cancel = 1
            Exit Sub
        End If
    End If
    
    '��������˹��̣������Ƿ����ļ�δ����
    If CurrentProject.ProjectType <> 0 Then
        If IsSaveRequired() Then
            Dim SaveRtn     As Integer                      '���淵��ֵ
            
            SaveRtn = mnuSave_Click()
            If SaveRtn = 2 Or SaveRtn = 3 Then              '�����������û�ȡ������
                Cancel = 1
                Exit Sub
            End If
        End If
    End If
    
    '�Ѵ�������������������β
    Me.Hide
    
    '�ָ��������໯
    SetWindowLongA Me.hwnd, GWL_WNDPROC, GetPropA(Me.hwnd, "PrevWndProc")
    
    '�رչ�����ʾ�ı�����
    Call DestroyToolTip
    
    '�رչܵ�
    Me.tmrCheckProcess.Enabled = False                      'ֹͣ��ʱ��
    If Not GdbPipe Is Nothing Then
        GdbPipe.StopRecvOutput
        GdbPipe.CloseDosIO
    End If
    
    '�ر����д���
    Dim CodeWindow  As Form
    IsExiting = True                                        '�����˳�״̬
    For Each CodeWindow In CodeWindows                      'ж�����д��봰��
        Unload CodeWindow
    Next CodeWindow
    Unload NewCreateWindow
    Unload frmControlBox
    Unload frmCreateOptions
    Unload frmCreate
    Unload frmProperties
    Unload frmSolutionExplorer
    Unload frmErrorList
    Unload frmOutput
    Unload frmBreakpoints
    Unload frmWatch
    Unload frmLocals
    Unload frmImmediate
    Unload frmCallStack
    Unload frmThreads
    Unload frmModules
    Unload frmMemory
    Unload frmRegisters
    Unload frmDisassembly
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    '�������ͻ������Ĵ�С
    Me.picToolBar.Width = Me.ScaleWidth
    Me.picClientArea.Move 0, Me.picToolBar.Top + Me.picToolBar.Height, Me.ScaleWidth, Me.ScaleHeight - Me.picClientArea.Top
    
    '������󻯵��Ӵ���Ĵ�С
    Dim Target  As Form
    Dim wp      As WINDOWPLACEMENT
    
    For Each Target In Forms
        If GetPropA(Target.hwnd, "Parent") = Me.picWindowClientArea.hwnd Then
            GetWindowPlacement Target.hwnd, wp
            If wp.ShowCmd = SW_MAXIMIZE Then
                ShowWindow Target.hwnd, SW_HIDE
                ShowWindow Target.hwnd, SW_MAXIMIZE
            End If
        End If
    Next Target
End Sub

Private Sub picToolBar_Click()
    Me.picToolBar.ZOrder
End Sub

Private Sub TabBar_TabClick(Frm As Form, Index As Integer)
    On Error Resume Next
    Frm.SetFocus                                                                    '����TabBar֮���ö�Ӧ�Ĵ��ڻ�ý���
    Frm.SyntaxEdit.SetFocus
End Sub

Private Sub TabBar_WindowDropIn(Frm As Form, Index As Integer)
    On Error Resume Next
    Frm.SetFocus                                                                    '�����Ͻ������ö�Ӧ�Ĵ��ڻ�ý���
    Frm.SyntaxEdit.SetFocus
End Sub

Private Sub TabBar_WindowDropOut(Frm As Form, Index As Integer)
    On Error Resume Next
    Frm.SetFocus                                                                    '�����ϳ�ȥ���ö�Ӧ�Ĵ��ڻ�ý���
    Frm.SyntaxEdit.SetFocus
End Sub

Private Sub tmrCheckProcess_Timer()
    On Error Resume Next
    Dim PipeOutput                  As String                                       '�ܵ��������
    Dim PipeOutputLine()            As String                                       '�ܵ������ÿһ��
    Dim SplitTmp                    As String                                       '�ַ����ָ��
    Dim ExitCode                    As Long                                         '�����˳���
    Dim NewCodeWindow               As frmCodeWindow                                '�´����Ĵ����
    Dim i                           As Long
    
    If Not ProcessExists(GdbPipe.hProcess) Then
        frmOutput.OutputLog "gdb����" & GdbPipe.dwProcessId & "(" & Hex(GdbPipe.dwProcessId) & ") " & "�����˳������Ա��Ƚ�����"
        frmOutput.OutputLog "����ͼǿ�ƽ������Խ���" & DebugProgramInfo.dwProcessId & "(" & Hex(DebugProgramInfo.dwProcessId) & ")"
        If TerminateProcess(DebugProgramInfo.hProcess, 0) = 0 Then
            frmOutput.OutputLog "��������" & DebugProgramInfo.dwProcessId & "(" & Hex(DebugProgramInfo.dwProcessId) & ") " & "ʧ�ܣ������н����ý��̡�"
        End If
        Call ProcessExitedHandler(0)
        Exit Sub
    End If
    If GdbPipe.DosOutput(PipeOutput, "(gdb) ") = 1 Then                             '��ȡgdb�Ƿ�������Ϣ
        PipeOutputLine = Split(PipeOutput, vbCrLf)                                      '�ָ��gdb�����ÿһ��
        For i = 0 To UBound(PipeOutputLine)                                             '���������ÿһ��
            PipeOutput = PipeOutputLine(i)
            If PipeOutput Like "*Breakpoint *,*at*" Then                                    '�ϵ�������Ϣ��Breakpoint *, * at *:*��
                Dim BreakpointIndex     As Long                                                 '��ǰ���еĶϵ����ţ�gdb��
                Dim bSourceFileFound    As Boolean                                              '�ܷ��ҵ���Ӧ�Ĵ����ļ�
                Dim SourceLn            As Long                                                 '��Ӧ�Ĵ����к�
                
                CurrState = 2                                                                   '���µ���״̬
                Call AdjustRuntimeMenu                                                          '���²˵�״̬
                SplitTmp = Split(PipeOutput, "Breakpoint ")(1)                                  '��Breakpoint [*, * at *:*]��
                BreakpointIndex = CLng(Split(SplitTmp, ", ")(0))                                '��[*], * at *:*��
                SplitTmp = Right(SplitTmp, Len(SplitTmp) - InStr(SplitTmp, " at "))             '��Breakpoint *, * at [*:*]��
                SourceLn = CLng(Right(SplitTmp, Len(SplitTmp) - InStrRev(SplitTmp, ":")))       '��*:[*]��
                
                '�л�����Ӧ�Ĵ����
                '�Է���һ����ʱ���ȡ�ϵ��Ӧ��gdb�ϵ��ʱ����������GdbBreakpoints��ӳ����ȱ©�������ᵼ�´���һ����Ч�Ĵ��봰��
                If BreakpointIndex <= UBound(GdbBreakpoints) Then
                    Set NewCodeWindow = ShowCodeWindow(GdbBreakpoints(BreakpointIndex).FileIndex)       '�ڴ������ʾ�ϵ��Ӧ��λ��
                    If NewCodeWindow Is Nothing Then
                        NoSkinMsgBox Lang_Main_Debug_OpenSourceFailure & CurrentProject.Files(GdbBreakpoints(BreakpointIndex).FileIndex).FilePath, _
                            vbExclamation, Lang_Msgbox_Error
                    Else
                        NewCodeWindow.SyntaxEdit.CurrPos.Row = SourceLn                                     '��ת����Ӧ�Ĵ�����
                        NewCodeWindow.BreakLine = SourceLn
                        NewCodeWindow.SyntaxEdit.SetFocus
                        NewCodeWindow.RedrawBreakpoints                                                     '�����ж��е�С��ͷ
                    End If
                End If
                
                '�ڡ������������Ӷϵ�������Ϣ
                frmOutput.OutputLog Lang_Main_Debug_BreakpointHit & ": " & _
                    CurrentProject.Files(GdbBreakpoints(BreakpointIndex).FileIndex).FilePath & ":" & SourceLn
                frmOutput.AddLineInfo False, CurrentProject.Files(GdbBreakpoints(BreakpointIndex).FileIndex).FilePath, SourceLn
                
                '��ȡ���ֵ�����Ϣ
                Call frmCallStack.GetCallStack                                                  '��ȡ���ö�ջ
                Call frmLocals.GetLocals                                                        '��ȡ���ر���
            '======================================================================================================================
            
            ElseIf PipeOutput Like "[[]Inferior * exited *" Then                            '�����˳���Ϣ��[Inferior * (process *) exited *]�����°�gdb��
                SplitTmp = Right(PipeOutput, Len(PipeOutput) - InStrRev(PipeOutput, " exited ") - 7)    '��[Inferior * (process *) exited ��*]����
                SplitTmp = Left(SplitTmp, InStrRev(SplitTmp, "]") - 1)                                  '����*��]��
                
                If SplitTmp = "normally" Then                                                   '������������������0
                    ExitCode = 0
                Else                                                                            '����Ͱѷ���ֵ���˽��ƣ�ת��ʮ����
                    ExitCode = CLng("&O" & Right(SplitTmp, Len(SplitTmp) - InStrRev(SplitTmp, " ")))
                End If
                Call ProcessExitedHandler(ExitCode)
            
            '======================================================================================================================
            ElseIf PipeOutput Like "Program exited *" Then                                  '�����˳���Ϣ��Program exited *�����ɰ�gdb��
                SplitTmp = Right(PipeOutput, Len(PipeOutput) - InStrRev(PipeOutput, " "))       '��Program exited with code [*.]����Program exited [normally.]��
                SplitTmp = Left(SplitTmp, Len(SplitTmp) - 1)                                    '��[*].����[normally].��
                
                If SplitTmp = "normally" Then                                                   '������������������0
                    ExitCode = 0
                Else                                                                            '����Ͱѷ���ֵ���˽��ƣ�ת��ʮ����
                    ExitCode = CLng("&O" & SplitTmp)
                End If
                Call ProcessExitedHandler(ExitCode)
                
            '======================================================================================================================
            ElseIf PipeOutput Like "Program received signal *" Then                         '�����׳��쳣 ��Program Received signal *��
                frmOutput.OutputLog PipeOutput
                CurrState = 2                                                                   '���µ���״̬
                Call AdjustRuntimeMenu                                                          '���²˵�״̬
                
                Dim rtnInfo     As CallStackInfoStruct
                
                frmMain.GdbPipe.ClearPipe                                                       '��չܵ��������
                frmMain.GdbPipe.DosInput "frame" & vbCrLf                                       '��gdb���ͻ�ǰ����λ������
                frmMain.GdbPipe.DosOutput PipeOutput, "(gdb) "                                  '��ȡgdb���
                
                PipeOutputLine = Split(PipeOutput, vbCrLf)                                      '�ָ��gdb�����ÿһ��
                rtnInfo = ParseCallStackString(PipeOutputLine(0))                               '�����������ȡ��ǰ����λ��
                
                '�л�����Ӧ�Ĵ����
                Set NewCodeWindow = ShowCodeWindow(, rtnInfo.File)                              '�ڴ������ʾ�ϵ��Ӧ��λ��
                If NewCodeWindow Is Nothing Then
                    NoSkinMsgBox Lang_Main_Debug_OpenSourceFailure & rtnInfo.File, _
                            vbExclamation, Lang_Msgbox_Error
                Else
                    NewCodeWindow.SyntaxEdit.CurrPos.Row = rtnInfo.Line                             '��ת����Ӧ�Ĵ�����
                    NewCodeWindow.BreakLine = rtnInfo.Line
                    NewCodeWindow.SyntaxEdit.SetFocus
                    NewCodeWindow.RedrawBreakpoints                                                 '�����ж��е�С��ͷ
                End If
                
                '�ڡ������������ӳ����ж���Ϣ
                frmOutput.OutputLog "�����ж��� " & rtnInfo.File & ":" & rtnInfo.Line & " (" & rtnInfo.Address & ")"        'todo: translate
                frmOutput.AddLineInfo False, rtnInfo.File, rtnInfo.Line
                
                '��ȡ���ֵ�����Ϣ
                Call frmCallStack.GetCallStack                                                  '��ȡ���ö�ջ
                Call frmLocals.GetLocals                                                        '��ȡ���ر���
            End If
        Next i
    Else
        Me.tmrCheckProcess.Enabled = False
        MsgBox "gdb BOOMED!"
    End If
End Sub

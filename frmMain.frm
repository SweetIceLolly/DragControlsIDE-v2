VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "CO7FCA~1.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "COE2B7~1.OCX"
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Image.ocx"
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
   Begin DragControlsIDE.DarkMenu DarkMenu 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      TabIndex        =   7
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
      MenuIcon_3      =   "frmMain.frx":1BD02
      SubMenuID_3_0   =   0
      MenuID_4        =   3
      MenuText_4      =   "���� (&S)           Ctrl+S"
      MenuVisible_4   =   -1  'True
      MenuIcon_4      =   "frmMain.frx":1BD22
      SubMenuID_4_0   =   0
      MenuID_5        =   4
      MenuText_5      =   "���Ϊ (&A)         Ctrl+Shift+S"
      MenuVisible_5   =   -1  'True
      MenuIcon_5      =   "frmMain.frx":1BD42
      SubMenuID_5_0   =   0
      MenuID_6        =   5
      MenuText_6      =   "-"
      MenuVisible_6   =   -1  'True
      MenuIcon_6      =   "frmMain.frx":1BD62
      SubMenuID_6_0   =   0
      MenuID_7        =   6
      MenuText_7      =   "�˳� (&E)"
      MenuVisible_7   =   -1  'True
      MenuIcon_7      =   "frmMain.frx":1BD82
      SubMenuID_7_0   =   0
      MenuID_8        =   7
      MenuText_8      =   "�༭"
      MenuVisible_8   =   -1  'True
      MenuIcon_8      =   "frmMain.frx":1BDA2
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
      MenuIcon_9      =   "frmMain.frx":1BDC2
      SubMenuID_9_0   =   0
      MenuID_10       =   9
      MenuText_10     =   "�ظ� (&R)           Ctrl+Y"
      MenuVisible_10  =   -1  'True
      MenuIcon_10     =   "frmMain.frx":1BDE2
      SubMenuID_10_0  =   0
      MenuID_11       =   10
      MenuText_11     =   "-"
      MenuVisible_11  =   -1  'True
      MenuIcon_11     =   "frmMain.frx":1BE02
      SubMenuID_11_0  =   0
      MenuID_12       =   11
      MenuText_12     =   "���� (&U)           Ctrl+X"
      MenuVisible_12  =   -1  'True
      MenuIcon_12     =   "frmMain.frx":1BE22
      SubMenuID_12_0  =   0
      MenuID_13       =   12
      MenuText_13     =   "���� (&C)           Ctrl+C"
      MenuVisible_13  =   -1  'True
      MenuIcon_13     =   "frmMain.frx":1BE42
      SubMenuID_13_0  =   0
      MenuID_14       =   13
      MenuText_14     =   "ճ�� (&P)           Ctrl+V"
      MenuVisible_14  =   -1  'True
      MenuIcon_14     =   "frmMain.frx":1BE62
      SubMenuID_14_0  =   0
      MenuID_15       =   14
      MenuText_15     =   "ȫѡ (&S)           Ctrl+A"
      MenuVisible_15  =   -1  'True
      MenuIcon_15     =   "frmMain.frx":1BE82
      SubMenuID_15_0  =   0
      MenuID_16       =   15
      MenuText_16     =   "ɾ���� (&D)         Ctrl+L"
      MenuVisible_16  =   -1  'True
      MenuIcon_16     =   "frmMain.frx":1BEA2
      SubMenuID_16_0  =   0
      MenuID_17       =   16
      MenuText_17     =   "-"
      MenuVisible_17  =   -1  'True
      MenuIcon_17     =   "frmMain.frx":1BEC2
      SubMenuID_17_0  =   0
      MenuID_18       =   17
      MenuText_18     =   "���� (&F)           Ctrl+F"
      MenuVisible_18  =   -1  'True
      MenuIcon_18     =   "frmMain.frx":1BEE2
      SubMenuID_18_0  =   0
      MenuID_19       =   18
      MenuText_19     =   "�滻 (&E)           Ctrl+H"
      MenuVisible_19  =   -1  'True
      MenuIcon_19     =   "frmMain.frx":1BF02
      SubMenuID_19_0  =   0
      MenuID_20       =   19
      MenuText_20     =   "-"
      MenuVisible_20  =   -1  'True
      MenuIcon_20     =   "frmMain.frx":1BF22
      SubMenuID_20_0  =   0
      MenuID_21       =   20
      MenuText_21     =   "�������� (&I)       Tab"
      MenuVisible_21  =   -1  'True
      MenuIcon_21     =   "frmMain.frx":1BF42
      SubMenuID_21_0  =   0
      MenuID_22       =   21
      MenuText_22     =   "�������� (&O)       Shift+Tab"
      MenuVisible_22  =   -1  'True
      MenuIcon_22     =   "frmMain.frx":1BF62
      SubMenuID_22_0  =   0
      MenuID_23       =   22
      MenuText_23     =   "-"
      MenuVisible_23  =   -1  'True
      MenuIcon_23     =   "frmMain.frx":1BF82
      SubMenuID_23_0  =   0
      MenuID_24       =   23
      MenuText_24     =   "���/�Ƴ��ϵ� (&B)  F9"
      MenuVisible_24  =   -1  'True
      MenuIcon_24     =   "frmMain.frx":1BFA2
      SubMenuID_24_0  =   0
      MenuID_25       =   24
      MenuText_25     =   "������жϵ� (&M)"
      MenuVisible_25  =   -1  'True
      MenuIcon_25     =   "frmMain.frx":1BFC2
      SubMenuID_25_0  =   0
      MenuID_26       =   25
      MenuText_26     =   "-"
      MenuVisible_26  =   -1  'True
      MenuIcon_26     =   "frmMain.frx":1BFE2
      SubMenuID_26_0  =   0
      MenuID_27       =   26
      MenuText_27     =   "��ת���� (&J)       Ctrl+G"
      MenuVisible_27  =   -1  'True
      MenuIcon_27     =   "frmMain.frx":1C002
      SubMenuID_27_0  =   0
      MenuID_28       =   27
      MenuText_28     =   "��ͼ"
      MenuVisible_28  =   -1  'True
      MenuIcon_28     =   "frmMain.frx":1C022
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
      MenuIcon_29     =   "frmMain.frx":1C042
      SubMenuID_29_0  =   0
      MenuID_30       =   29
      MenuText_30     =   "�ؼ��� (&C)"
      MenuCheckBox_30 =   -1  'True
      MenuVisible_30  =   -1  'True
      MenuIcon_30     =   "frmMain.frx":1C062
      SubMenuID_30_0  =   0
      MenuID_31       =   30
      MenuText_31     =   "���� (&P)           F4"
      MenuCheckBox_31 =   -1  'True
      MenuVisible_31  =   -1  'True
      MenuIcon_31     =   "frmMain.frx":1C082
      SubMenuID_31_0  =   0
      MenuID_32       =   31
      MenuText_32     =   "������Դ������ (&M)"
      MenuCheckBox_32 =   -1  'True
      MenuVisible_32  =   -1  'True
      MenuIcon_32     =   "frmMain.frx":1C0A2
      SubMenuID_32_0  =   0
      MenuID_33       =   32
      MenuText_33     =   "�����б� (&E)       Ctrl+E"
      MenuCheckBox_33 =   -1  'True
      MenuVisible_33  =   -1  'True
      MenuIcon_33     =   "frmMain.frx":1C0C2
      SubMenuID_33_0  =   0
      MenuID_34       =   33
      MenuText_34     =   "��� (&O)           Ctrl+Alt+O"
      MenuCheckBox_34 =   -1  'True
      MenuVisible_34  =   -1  'True
      MenuIcon_34     =   "frmMain.frx":1C0E2
      SubMenuID_34_0  =   0
      MenuID_35       =   34
      MenuText_35     =   "����"
      MenuVisible_35  =   -1  'True
      MenuIcon_35     =   "frmMain.frx":1C102
      SUBMENU_ITEM_COUNT_35=   2
      SubMenuID_35_0  =   0
      SubMenuText_35_1=   "���ɴ����ļ� (&C)"
      SubMenuID_35_1  =   36
      SubMenuText_35_2=   "���ɿ�ִ���ļ� (&E) Ctrl+F5"
      SubMenuID_35_2  =   37
      MenuID_36       =   35
      MenuText_36     =   "���ɴ����ļ� (&C)"
      MenuVisible_36  =   -1  'True
      MenuIcon_36     =   "frmMain.frx":1C122
      SubMenuID_36_0  =   0
      MenuID_37       =   36
      MenuText_37     =   "���ɿ�ִ���ļ� (&E) Ctrl+F5"
      MenuVisible_37  =   -1  'True
      MenuIcon_37     =   "frmMain.frx":1C142
      SubMenuID_37_0  =   0
      MenuID_38       =   37
      MenuText_38     =   "����"
      MenuVisible_38  =   -1  'True
      MenuIcon_38     =   "frmMain.frx":1C162
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
      MenuIcon_39     =   "frmMain.frx":1C182
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
      MenuIcon_40     =   "frmMain.frx":1C1A2
      SubMenuID_40_0  =   0
      MenuID_41       =   40
      MenuText_41     =   "-"
      MenuVisible_41  =   -1  'True
      MenuIcon_41     =   "frmMain.frx":1C1C2
      SubMenuID_41_0  =   0
      MenuID_42       =   41
      MenuText_42     =   "���Ӵ��� (&W)       Ctrl+Alt+W"
      MenuCheckBox_42 =   -1  'True
      MenuVisible_42  =   -1  'True
      MenuIcon_42     =   "frmMain.frx":1C1E2
      SubMenuID_42_0  =   0
      MenuID_43       =   42
      MenuText_43     =   "���� (&L)           Ctrl+Alt+L"
      MenuCheckBox_43 =   -1  'True
      MenuVisible_43  =   -1  'True
      MenuIcon_43     =   "frmMain.frx":1C202
      SubMenuID_43_0  =   0
      MenuID_44       =   43
      MenuText_44     =   "�������� (&I)       Ctrl+Alt+I"
      MenuCheckBox_44 =   -1  'True
      MenuVisible_44  =   -1  'True
      MenuIcon_44     =   "frmMain.frx":1C222
      SubMenuID_44_0  =   0
      MenuID_45       =   44
      MenuText_45     =   "-"
      MenuVisible_45  =   -1  'True
      MenuIcon_45     =   "frmMain.frx":1C242
      SubMenuID_45_0  =   0
      MenuID_46       =   45
      MenuText_46     =   "���ö�ջ (&C)       Ctrl+Alt+C"
      MenuCheckBox_46 =   -1  'True
      MenuVisible_46  =   -1  'True
      MenuIcon_46     =   "frmMain.frx":1C262
      SubMenuID_46_0  =   0
      MenuID_47       =   46
      MenuText_47     =   "�߳� (&T)           Ctrl+Alt+T"
      MenuCheckBox_47 =   -1  'True
      MenuVisible_47  =   -1  'True
      MenuIcon_47     =   "frmMain.frx":1C282
      SubMenuID_47_0  =   0
      MenuID_48       =   47
      MenuText_48     =   "ģ�� (&M)           Ctrl+Alt+M"
      MenuCheckBox_48 =   -1  'True
      MenuVisible_48  =   -1  'True
      MenuIcon_48     =   "frmMain.frx":1C2A2
      SubMenuID_48_0  =   0
      MenuID_49       =   48
      MenuText_49     =   "-"
      MenuVisible_49  =   -1  'True
      MenuIcon_49     =   "frmMain.frx":1C2C2
      SubMenuID_49_0  =   0
      MenuID_50       =   49
      MenuText_50     =   "�ڴ� (&E)           Ctrl+Alt+E"
      MenuCheckBox_50 =   -1  'True
      MenuVisible_50  =   -1  'True
      MenuIcon_50     =   "frmMain.frx":1C2E2
      SubMenuID_50_0  =   0
      MenuID_51       =   50
      MenuText_51     =   "�Ĵ��� (&R)         Ctrl+Alt+R"
      MenuCheckBox_51 =   -1  'True
      MenuVisible_51  =   -1  'True
      MenuIcon_51     =   "frmMain.frx":1C302
      SubMenuID_51_0  =   0
      MenuID_52       =   51
      MenuText_52     =   "����� (&D)         Ctrl+Alt+D"
      MenuCheckBox_52 =   -1  'True
      MenuVisible_52  =   -1  'True
      MenuIcon_52     =   "frmMain.frx":1C322
      SubMenuID_52_0  =   0
      MenuID_53       =   52
      MenuText_53     =   "���� (&R)           F5"
      MenuVisible_53  =   -1  'True
      MenuIcon_53     =   "frmMain.frx":1C342
      SubMenuID_53_0  =   0
      MenuID_54       =   53
      MenuText_54     =   "�ж� (&B)           Ctrl+Alt+Break"
      MenuVisible_54  =   -1  'True
      MenuIcon_54     =   "frmMain.frx":1C362
      SubMenuID_54_0  =   0
      MenuID_55       =   54
      MenuText_55     =   "ֹͣ (&E)           Shift+F5"
      MenuVisible_55  =   -1  'True
      MenuIcon_55     =   "frmMain.frx":1C382
      SubMenuID_55_0  =   0
      MenuID_56       =   55
      MenuText_56     =   "�������� (&S)       Ctrl+Shift+F5"
      MenuVisible_56  =   -1  'True
      MenuIcon_56     =   "frmMain.frx":1C3A2
      SubMenuID_56_0  =   0
      MenuID_57       =   56
      MenuText_57     =   "-"
      MenuVisible_57  =   -1  'True
      MenuIcon_57     =   "frmMain.frx":1C3C2
      SubMenuID_57_0  =   0
      MenuID_58       =   57
      MenuText_58     =   "�����ִ��         F11"
      MenuVisible_58  =   -1  'True
      MenuIcon_58     =   "frmMain.frx":1C3E2
      SubMenuID_58_0  =   0
      MenuID_59       =   58
      MenuText_59     =   "�����ִ��         F10"
      MenuVisible_59  =   -1  'True
      MenuIcon_59     =   "frmMain.frx":1C402
      SubMenuID_59_0  =   0
      MenuID_60       =   59
      MenuText_60     =   "ִ�е�����         Shift+F11"
      MenuVisible_60  =   -1  'True
      MenuIcon_60     =   "frmMain.frx":1C422
      SubMenuID_60_0  =   0
      MenuID_61       =   60
      MenuText_61     =   "����"
      MenuVisible_61  =   -1  'True
      MenuIcon_61     =   "frmMain.frx":1C442
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
      MenuIcon_62     =   "frmMain.frx":1C462
      SubMenuID_62_0  =   0
      MenuID_63       =   62
      MenuText_63     =   "��Ϣ���� (&M)"
      MenuVisible_63  =   -1  'True
      MenuIcon_63     =   "frmMain.frx":1C482
      SubMenuID_63_0  =   0
      MenuID_64       =   63
      MenuText_64     =   "���� (&P)"
      MenuVisible_64  =   -1  'True
      MenuIcon_64     =   "frmMain.frx":1C4A2
      SubMenuID_64_0  =   0
      MenuID_65       =   64
      MenuText_65     =   "-"
      MenuVisible_65  =   -1  'True
      MenuIcon_65     =   "frmMain.frx":1C4C2
      SubMenuID_65_0  =   0
      MenuID_66       =   65
      MenuText_66     =   "���� (&O)"
      MenuVisible_66  =   -1  'True
      MenuIcon_66     =   "frmMain.frx":1C4E2
      SubMenuID_66_0  =   0
      MenuID_67       =   66
      MenuText_67     =   "����"
      MenuVisible_67  =   -1  'True
      MenuIcon_67     =   "frmMain.frx":1C502
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
      MenuIcon_68     =   "frmMain.frx":1C522
      SubMenuID_68_0  =   0
      MenuID_69       =   68
      MenuText_69     =   "ʾ������ (&E)"
      MenuVisible_69  =   -1  'True
      MenuIcon_69     =   "frmMain.frx":1C542
      SubMenuID_69_0  =   0
      MenuID_70       =   69
      MenuText_70     =   "�����Ͽؼ��� (&A) Ctrl+F1"
      MenuVisible_70  =   -1  'True
      MenuIcon_70     =   "frmMain.frx":1C562
      SubMenuID_70_0  =   0
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
      TabIndex        =   6
      Top             =   804
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
      Height          =   5625
      Left            =   0
      ScaleHeight     =   5625
      ScaleWidth      =   16845
      TabIndex        =   0
      Top             =   1185
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
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   5655
      End
      Begin DragControlsIDE.DarkButton cmdOpenProject 
         Height          =   615
         Left            =   1680
         TabIndex        =   4
         Top             =   3000
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft YaHei UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�򿪹���..."
         HasBorder       =   0   'False
         Alignment       =   0
      End
      Begin DragControlsIDE.DarkButton cmdNewConsoleProgram 
         Height          =   615
         Left            =   1680
         TabIndex        =   2
         Top             =   1560
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft YaHei UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�½�����̨����"
         HasBorder       =   0   'False
         Alignment       =   0
      End
      Begin DragControlsIDE.DarkButton cmdNewPlainCpp 
         Height          =   615
         Left            =   1680
         TabIndex        =   3
         Top             =   2280
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft YaHei UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�½��հ�C++����"
         HasBorder       =   0   'False
         Alignment       =   0
      End
      Begin DragControlsIDE.DarkButton cmdNewWindowProgram 
         Height          =   615
         Left            =   1680
         TabIndex        =   1
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft YaHei UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�½����ڳ���"
         HasBorder       =   0   'False
         Alignment       =   0
      End
      Begin ImageX.aicAlphaImage imgNewWindowProgram 
         Height          =   645
         Left            =   600
         Top             =   840
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1138
         Image           =   "frmMain.frx":1C582
         Props           =   5
      End
      Begin VB.Label labTip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEDB1A&
         Height          =   330
         Index           =   3
         Left            =   600
         TabIndex        =   11
         Top             =   4200
         Width           =   480
      End
      Begin VB.Label labTip 
         Appearance      =   0  'Flat
         BackColor       =   &H00CEDB1A&
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFB00A&
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   10
         Top             =   4200
         Width           =   90
      End
      Begin ImageX.aicAlphaImage imgOpenProject 
         Height          =   645
         Left            =   600
         Top             =   3000
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         Image           =   "frmMain.frx":1C6B4
         Props           =   5
      End
      Begin ImageX.aicAlphaImage imgNewPlainCpp 
         Height          =   615
         Left            =   600
         Top             =   2280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Image           =   "frmMain.frx":1C8FB
         Props           =   5
      End
      Begin ImageX.aicAlphaImage imgNewConsoleProgram 
         Height          =   645
         Left            =   600
         Top             =   1560
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         Image           =   "frmMain.frx":1CC3F
         Props           =   5
      End
      Begin VB.Label labTip 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFB00A&
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFB00A&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   90
      End
      Begin VB.Label labTip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFB00A&
         Height          =   330
         Index           =   1
         Left            =   600
         TabIndex        =   8
         Top             =   240
         Width           =   480
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
      TabIndex        =   5
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
      BindCaption     =   -1  'True
      Picture         =   "frmMain.frx":1CDBF
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
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public ProjectType  As Integer                                                                                          '��������
'ֵ     ����
'0      δ�������̣�������������
'1      ���ڳ���
'2      ����̨����
'3      �հ�C++����

'����: ������������
Private Sub HideStartupPage()
    Dim i           As Integer
    
    For i = 0 To 3
        Me.labTip(i).Visible = False
    Next i
    Me.imgNewConsoleProgram.Visible = False
    Me.imgNewPlainCpp.Visible = False
    Me.imgNewWindowProgram.Visible = False
    Me.imgOpenProject.Visible = False
    Me.cmdNewConsoleProgram.Visible = False
    Me.cmdNewPlainCpp.Visible = False
    Me.cmdNewWindowProgram.Visible = False
    Me.cmdOpenProject.Visible = False
    
    Me.DarkMenu.MenuEnabled(3) = True                                                                                   '����
    Me.DarkMenu.MenuEnabled(4) = True                                                                                   '���Ϊ
    Me.DarkMenu.MenuEnabled(7) = True                                                                                   '�༭
    Me.DarkMenu.MenuEnabled(27) = True                                                                                  '��ͼ
    Me.DarkMenu.MenuEnabled(34) = True                                                                                  '����
    Me.DarkMenu.MenuEnabled(37) = True                                                                                  '����
End Sub

Private Sub cmdNewConsoleProgram_Click()
    Call HideStartupPage
End Sub

Private Sub cmdNewPlainCpp_Click()
    On Error Resume Next
    
    ProjectType = 3                                                                                                     '���ù�������
    Call HideStartupPage                                                                                                '������������
    
    SetPropA frmCodeWindow.hWnd, "Parent", Me.picWindowClientArea.hWnd                                                  '��¼���봰�ڵ�ĸ����, ���������С֮��
    SetParent frmCodeWindow.hWnd, Me.picWindowClientArea.hWnd                                                           '���ô��봰�ڵ�ĸ����
    frmCodeWindow.Move 0, 0, (Me.picClientArea.ScaleWidth - 250) / 3 * 2, (Me.picClientArea.ScaleHeight - 120) / 3 * 2  '����������С
    
    '���ò����õĲ˵���
    Me.DarkMenu.MenuEnabled(29) = False                                                                                 '�ؼ���
    Me.DarkMenu.MenuEnabled(30) = False                                                                                 '����
    
    '��ʾ��Ҫ��Pane
    Me.DockingPane.ShowPane 3                                                                                           '������Դ������
    Me.DockingPane.ShowPane 5                                                                                           '���
    
    '���±���
    Me.Caption = "�¹��� - �Ͽؼ���"
    
    '�ô�����ý���
    Me.picWindowClientArea.Visible = True
    frmCodeWindow.Show
    frmCodeWindow.SyntaxEdit.SetFocus
End Sub

Private Sub cmdNewWindowProgram_Click()
    Call HideStartupPage
End Sub

Private Sub cmdOpenProject_Click()
    Call HideStartupPage
End Sub

Private Sub DarkMenu_MenuItemClicked(MenuID As Integer)
    Select Case MenuID
        
    End Select
End Sub

Private Sub DockingPane_Resize()
    If ProjectType <> 0 Then                                                            '�����������������Ļ��͵������ڻ�ͻ�����С
        Dim cLeft   As Long, cRight   As Long, cTop   As Long, cBottom   As Long
        
        Me.DockingPane.GetClientRect cLeft, cTop, cRight, cBottom
        Me.picWindowClientArea.Move cLeft, cTop, cRight - cLeft, cBottom - cTop
        
        Call Form_Resize                                                                    '������ڿͻ�����������󻯵Ĵ��ڣ������С���е���
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    '����LOGO
    frmStartupLogo.Show
    
    '�����ַ�����Դ
    If Not LoadLanguage(1001) Then
        MsgBox "�����ַ�����Դʧ�ܣ�" & Err.Number & ": " & Err.Description, vbCritical, "����"
    End If
    
    '�������ͻ�����
    Me.picClientArea.Height = Me.ScaleHeight - Me.picClientArea.Top                                                     '�����ͻ������Ĵ�С
    Me.picWindowClientArea.BackColor = Me.BackColor                                                                     '���ڿͻ�������ɫ
    
    '����������
    Dim ClientHeight        As Integer, ClientWidth             As Integer
    Dim i                   As Integer
    
    Me.DockingPane.AttachToWindow Me.picClientArea.hWnd                                                                 '�󶨹�����
    ClientHeight = Me.picClientArea.ScaleHeight / Screen.TwipsPerPixelY
    ClientWidth = Me.picClientArea.ScaleWidth / Screen.TwipsPerPixelX
    Me.DockingPane.CreatePane 1, 100, ClientHeight, DockLeftOf                                                          '�ؼ���
    Me.DockingPane(1).Handle = frmControlBox.hWnd
    Me.DockingPane(1).Title = "�ؼ���"
    Me.DockingPane.CreatePane 2, 250, ClientHeight / 2, DockRightOf                                                     '����
    Me.DockingPane(2).Handle = frmProperties.hWnd
    Me.DockingPane(2).Title = "����"
    Me.DockingPane.CreatePane 3, 250, ClientHeight / 2, DockRightOf                                                     '������Դ������
    Me.DockingPane(3).Handle = frmSolutionExplorer.hWnd
    Me.DockingPane(3).Title = "������Դ������"
    Me.DockingPane.CreatePane 4, ClientWidth / 2, 120, DockBottomOf Or DockLeftOf                                       '�����б�
    Me.DockingPane(4).Handle = frmErrorList.hWnd
    Me.DockingPane(4).Title = "�����б�"
    Me.DockingPane.CreatePane 5, ClientWidth / 2, 120, DockBottomOf Or DockRightOf                                      '���
    Me.DockingPane(5).Handle = frmOutput.hWnd
    Me.DockingPane(5).Title = "���"
    Me.DockingPane.CreatePane 6, ClientWidth / 2, 120, DockBottomOf Or DockRightOf                                      '�ϵ��б�
    Me.DockingPane(6).Handle = frmBreakpoints.hWnd
    Me.DockingPane(6).Title = "�ϵ��б�"
    Me.DockingPane.CreatePane 7, ClientWidth / 2, 120, DockBottomOf                                                     '���Ӵ���
    Me.DockingPane(7).Handle = frmWatch.hWnd
    Me.DockingPane(7).Title = "���Ӵ���"
    Me.DockingPane.CreatePane 8, ClientWidth / 2, 120, DockBottomOf                                                     '����
    Me.DockingPane(8).Handle = frmLocals.hWnd
    Me.DockingPane(8).Title = "����"
    Me.DockingPane.CreatePane 9, ClientWidth / 2, 120, DockBottomOf                                                     '��������
    Me.DockingPane(9).Handle = frmImmediate.hWnd
    Me.DockingPane(9).Title = "��������"
    Me.DockingPane.CreatePane 10, ClientWidth / 2, 120, DockBottomOf                                                    '���ö�ջ
    Me.DockingPane(10).Handle = frmCallStack.hWnd
    Me.DockingPane(10).Title = "���ö�ջ"
    Me.DockingPane.CreatePane 11, ClientWidth / 2, 120, DockBottomOf                                                    '�߳�
    Me.DockingPane(11).Handle = frmThreads.hWnd
    Me.DockingPane(11).Title = "�߳�"
    Me.DockingPane.CreatePane 12, ClientWidth / 2, 120, DockBottomOf                                                    'ģ��
    Me.DockingPane(12).Handle = frmModules.hWnd
    Me.DockingPane(12).Title = "ģ��"
    Me.DockingPane.CreatePane 13, ClientWidth / 2, 250, DockBottomOf                                                    '�ڴ�
    Me.DockingPane(13).Handle = frmMemory.hWnd
    Me.DockingPane(13).Title = "�ڴ�"
    Me.DockingPane.CreatePane 14, ClientWidth / 2, 250, DockBottomOf                                                    '�Ĵ���
    Me.DockingPane(14).Handle = frmRegisters.hWnd
    Me.DockingPane(14).Title = "�Ĵ���"
    Me.DockingPane.CreatePane 15, ClientWidth / 2, 250, DockBottomOf                                                    '�����
    Me.DockingPane(15).Handle = frmDisassembly.hWnd
    Me.DockingPane(15).Title = "�����"
    For i = 1 To 15                                                                                                     '�������е�Pane
        Me.DockingPane(i).Close
    Next i
    
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
    Me.SkinFramework.LoadSkin "Skin.cjstyles", "NormalBlue.ini"
    Me.SkinFramework.ApplyWindow Me.hWnd
    
    '���ò���Ҫ�Ĳ˵�
    Me.DarkMenu.MenuEnabled(3) = False                                                                                  '����
    Me.DarkMenu.MenuEnabled(4) = False                                                                                  '���Ϊ
    Me.DarkMenu.MenuEnabled(7) = False                                                                                  '�༭
    Me.DarkMenu.MenuEnabled(27) = False                                                                                 '��ͼ
    Me.DarkMenu.MenuEnabled(34) = False                                                                                 '����
    Me.DarkMenu.MenuEnabled(37) = False                                                                                 '����
    
    '���ô������໯������������⼰�����������Ҽ��ر�
    SetPropA Me.hWnd, "PrevWndProc", SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf MainWindowMaximizeCloseFixProc)
    
    'ж��LOGO
    Unload frmStartupLogo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '�ָ��������໯
    SetWindowLongA Me.hWnd, GWL_WNDPROC, GetPropA(hWnd, "PrevWndProc")
    
    '�ر����д���
    Unload frmCodeWindow
    Unload frmControlBox
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
    Me.picClientArea.Height = Me.ScaleHeight - Me.picClientArea.Top
    
    '������󻯵��Ӵ���Ĵ�С
    Dim Target  As Form
    Dim wp      As WINDOWPLACEMENT
    
    For Each Target In Forms
        If GetPropA(Target.hWnd, "Parent") = Me.picWindowClientArea.hWnd Then
            GetWindowPlacement Target.hWnd, wp
            If wp.ShowCmd = SW_MAXIMIZE Then
                ShowWindow Target.hWnd, SW_HIDE
                ShowWindow Target.hWnd, SW_MAXIMIZE
            End If
        End If
    Next Target
End Sub

Private Sub imgNewConsoleProgram_Click(ByVal Button As Integer)
    Call cmdNewConsoleProgram_Click
End Sub

Private Sub imgNewPlainCpp_Click(ByVal Button As Integer)
    Call cmdNewPlainCpp_Click
End Sub

Private Sub imgNewWindowProgram_Click(ByVal Button As Integer)
    Call cmdNewWindowProgram_Click
End Sub

Private Sub imgOpenProject_Click(ByVal Button As Integer)
    Call cmdOpenProject_Click
End Sub

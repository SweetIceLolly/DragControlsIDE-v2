VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "DOCKIN~1.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "SKINFR~1.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "�Ͽؼ���"
   ClientHeight    =   7908
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16848
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Microsoft YaHei UI"
      Size            =   8.4
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
   ScaleHeight     =   7908
   ScaleWidth      =   16848
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer tmrCheckProcess 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   13440
      Top             =   7320
   End
   Begin DragControlsIDE.DarkMenu DarkMenu 
      Align           =   1  'Align Top
      Height          =   312
      Left            =   0
      TabIndex        =   3
      Top             =   492
      Width           =   16848
      _extentx        =   29718
      _extenty        =   550
      font            =   "frmMain.frx":1BCC2
      menu_item_count =   70
      menuid_1        =   0
      menutext_1      =   "�ļ�"
      menuvisible_1   =   -1  'True
      menuicon_1      =   "frmMain.frx":1BCF6
      submenu_item_count_1=   6
      submenu_item_count_1=   6
      submenuid_1_0   =   0
      submenutext_1_1 =   "�½���Ŀ (&N)       Ctrl+N"
      submenuid_1_1   =   2
      submenutext_1_2 =   "������Ŀ (&O)       Ctrl+O"
      submenuid_1_2   =   3
      submenutext_1_3 =   "���� (&S)           Ctrl+S"
      submenuid_1_3   =   4
      submenutext_1_4 =   "���Ϊ (&A)         Ctrl+Shift+S"
      submenuid_1_4   =   5
      submenutext_1_5 =   "-"
      submenuid_1_5   =   6
      submenutext_1_6 =   "�˳� (&E)"
      submenuid_1_6   =   7
      menuid_2        =   1
      menutext_2      =   "�½���Ŀ (&N)       Ctrl+N"
      menuvisible_2   =   -1  'True
      menuicon_2      =   "frmMain.frx":1BD16
      submenuid_2_0   =   0
      menuid_3        =   2
      menutext_3      =   "������Ŀ (&O)       Ctrl+O"
      menuvisible_3   =   -1  'True
      menuicon_3      =   "frmMain.frx":1BE45
      submenuid_3_0   =   0
      menuid_4        =   3
      menutext_4      =   "���� (&S)           Ctrl+S"
      menuvisible_4   =   -1  'True
      menuicon_4      =   "frmMain.frx":1C02F
      submenuid_4_0   =   0
      menuid_5        =   4
      menutext_5      =   "���Ϊ (&A)         Ctrl+Shift+S"
      menuvisible_5   =   -1  'True
      menuicon_5      =   "frmMain.frx":1C13D
      submenuid_5_0   =   0
      menuid_6        =   5
      menutext_6      =   "-"
      menuvisible_6   =   -1  'True
      menuicon_6      =   "frmMain.frx":1C2FC
      submenuid_6_0   =   0
      menuid_7        =   6
      menutext_7      =   "�˳� (&E)"
      menuvisible_7   =   -1  'True
      menuicon_7      =   "frmMain.frx":1C31C
      submenuid_7_0   =   0
      menuid_8        =   7
      menutext_8      =   "�༭"
      menuvisible_8   =   -1  'True
      menuicon_8      =   "frmMain.frx":1C33C
      submenu_item_count_8=   19
      submenu_item_count_8=   19
      submenuid_8_0   =   0
      submenutext_8_1 =   "���� (&U)           Ctrl+Z"
      submenuid_8_1   =   9
      submenutext_8_2 =   "�ظ� (&R)           Ctrl+Y"
      submenuid_8_2   =   10
      submenutext_8_3 =   "-"
      submenuid_8_3   =   11
      submenutext_8_4 =   "���� (&U)           Ctrl+X"
      submenuid_8_4   =   12
      submenutext_8_5 =   "���� (&C)           Ctrl+C"
      submenuid_8_5   =   13
      submenutext_8_6 =   "ճ�� (&P)           Ctrl+V"
      submenuid_8_6   =   14
      submenutext_8_7 =   "ȫѡ (&S)           Ctrl+A"
      submenuid_8_7   =   15
      submenutext_8_8 =   "ɾ���� (&D)         Ctrl+L"
      submenuid_8_8   =   16
      submenutext_8_9 =   "-"
      submenuid_8_9   =   17
      submenutext_8_10=   "���� (&F)           Ctrl+F"
      submenuid_8_10  =   18
      submenutext_8_11=   "�滻 (&E)           Ctrl+H"
      submenuid_8_11  =   19
      submenutext_8_12=   "-"
      submenuid_8_12  =   20
      submenutext_8_13=   "�������� (&I)       Tab"
      submenuid_8_13  =   21
      submenutext_8_14=   "�������� (&O)       Shift+Tab"
      submenuid_8_14  =   22
      submenutext_8_15=   "-"
      submenuid_8_15  =   23
      submenutext_8_16=   "���/�Ƴ��ϵ� (&B)  F9"
      submenuid_8_16  =   24
      submenutext_8_17=   "������жϵ� (&M)"
      submenuid_8_17  =   25
      submenutext_8_18=   "-"
      submenuid_8_18  =   26
      submenutext_8_19=   "��ת���� (&J)       Ctrl+G"
      submenuid_8_19  =   27
      menuid_9        =   8
      menutext_9      =   "���� (&U)           Ctrl+Z"
      menuvisible_9   =   -1  'True
      menuicon_9      =   "frmMain.frx":1C35C
      submenuid_9_0   =   0
      menuid_10       =   9
      menutext_10     =   "�ظ� (&R)           Ctrl+Y"
      menuvisible_10  =   -1  'True
      menuicon_10     =   "frmMain.frx":1C52D
      submenuid_10_0  =   0
      menuid_11       =   10
      menutext_11     =   "-"
      menuvisible_11  =   -1  'True
      menuicon_11     =   "frmMain.frx":1C73C
      submenuid_11_0  =   0
      menuid_12       =   11
      menutext_12     =   "���� (&U)           Ctrl+X"
      menuvisible_12  =   -1  'True
      menuicon_12     =   "frmMain.frx":1C75C
      submenuid_12_0  =   0
      menuid_13       =   12
      menutext_13     =   "���� (&C)           Ctrl+C"
      menuvisible_13  =   -1  'True
      menuicon_13     =   "frmMain.frx":1C8BC
      submenuid_13_0  =   0
      menuid_14       =   13
      menutext_14     =   "ճ�� (&P)           Ctrl+V"
      menuvisible_14  =   -1  'True
      menuicon_14     =   "frmMain.frx":1CA17
      submenuid_14_0  =   0
      menuid_15       =   14
      menutext_15     =   "ȫѡ (&S)           Ctrl+A"
      menuvisible_15  =   -1  'True
      menuicon_15     =   "frmMain.frx":1CB6F
      submenuid_15_0  =   0
      menuid_16       =   15
      menutext_16     =   "ɾ���� (&D)         Ctrl+L"
      menuvisible_16  =   -1  'True
      menuicon_16     =   "frmMain.frx":1CC6E
      submenuid_16_0  =   0
      menuid_17       =   16
      menutext_17     =   "-"
      menuvisible_17  =   -1  'True
      menuicon_17     =   "frmMain.frx":1CC8E
      submenuid_17_0  =   0
      menuid_18       =   17
      menutext_18     =   "���� (&F)           Ctrl+F"
      menuvisible_18  =   -1  'True
      menuicon_18     =   "frmMain.frx":1CCAE
      submenuid_18_0  =   0
      menuid_19       =   18
      menutext_19     =   "�滻 (&E)           Ctrl+H"
      menuvisible_19  =   -1  'True
      menuicon_19     =   "frmMain.frx":1CDC1
      submenuid_19_0  =   0
      menuid_20       =   19
      menutext_20     =   "-"
      menuvisible_20  =   -1  'True
      menuicon_20     =   "frmMain.frx":1CF25
      submenuid_20_0  =   0
      menuid_21       =   20
      menutext_21     =   "�������� (&I)       Tab"
      menuvisible_21  =   -1  'True
      menuicon_21     =   "frmMain.frx":1CF45
      submenuid_21_0  =   0
      menuid_22       =   21
      menutext_22     =   "�������� (&O)       Shift+Tab"
      menuvisible_22  =   -1  'True
      menuicon_22     =   "frmMain.frx":1CF65
      submenuid_22_0  =   0
      menuid_23       =   22
      menutext_23     =   "-"
      menuvisible_23  =   -1  'True
      menuicon_23     =   "frmMain.frx":1CF85
      submenuid_23_0  =   0
      menuid_24       =   23
      menutext_24     =   "���/�Ƴ��ϵ� (&B)  F9"
      menuvisible_24  =   -1  'True
      menuicon_24     =   "frmMain.frx":1CFA5
      submenuid_24_0  =   0
      menuid_25       =   24
      menutext_25     =   "������жϵ� (&M)"
      menuvisible_25  =   -1  'True
      menuicon_25     =   "frmMain.frx":1CFC5
      submenuid_25_0  =   0
      menuid_26       =   25
      menutext_26     =   "-"
      menuvisible_26  =   -1  'True
      menuicon_26     =   "frmMain.frx":1CFE5
      submenuid_26_0  =   0
      menuid_27       =   26
      menutext_27     =   "��ת���� (&J)       Ctrl+G"
      menuvisible_27  =   -1  'True
      menuicon_27     =   "frmMain.frx":1D005
      submenuid_27_0  =   0
      menuid_28       =   27
      menutext_28     =   "��ͼ"
      menuvisible_28  =   -1  'True
      menuicon_28     =   "frmMain.frx":1D025
      submenu_item_count_28=   6
      submenu_item_count_28=   6
      submenuid_28_0  =   0
      submenutext_28_1=   "������ (&T)"
      submenuid_28_1  =   29
      submenutext_28_2=   "�ؼ��� (&C)"
      submenuid_28_2  =   30
      submenutext_28_3=   "���� (&P)           F4"
      submenuid_28_3  =   31
      submenutext_28_4=   "������Դ������ (&M)"
      submenuid_28_4  =   32
      submenutext_28_5=   "�����б� (&E)       Ctrl+E"
      submenuid_28_5  =   33
      submenutext_28_6=   "��� (&O)           Ctrl+Alt+O"
      submenuid_28_6  =   34
      menuid_29       =   28
      menutext_29     =   "������ (&T)"
      menucheckbox_29 =   -1  'True
      menuvisible_29  =   -1  'True
      menuicon_29     =   "frmMain.frx":1D045
      submenuid_29_0  =   0
      menuid_30       =   29
      menutext_30     =   "�ؼ��� (&C)"
      menucheckbox_30 =   -1  'True
      menuvisible_30  =   -1  'True
      menuicon_30     =   "frmMain.frx":1D10A
      submenuid_30_0  =   0
      menuid_31       =   30
      menutext_31     =   "���� (&P)           F4"
      menucheckbox_31 =   -1  'True
      menuvisible_31  =   -1  'True
      menuicon_31     =   "frmMain.frx":1D1BD
      submenuid_31_0  =   0
      menuid_32       =   31
      menutext_32     =   "������Դ������ (&M)"
      menucheckbox_32 =   -1  'True
      menuvisible_32  =   -1  'True
      menuicon_32     =   "frmMain.frx":1D349
      submenuid_32_0  =   0
      menuid_33       =   32
      menutext_33     =   "�����б� (&E)       Ctrl+E"
      menucheckbox_33 =   -1  'True
      menuvisible_33  =   -1  'True
      menuicon_33     =   "frmMain.frx":1D369
      submenuid_33_0  =   0
      menuid_34       =   33
      menutext_34     =   "��� (&O)           Ctrl+Alt+O"
      menucheckbox_34 =   -1  'True
      menuvisible_34  =   -1  'True
      menuicon_34     =   "frmMain.frx":1D517
      submenuid_34_0  =   0
      menuid_35       =   34
      menutext_35     =   "����"
      menuvisible_35  =   -1  'True
      menuicon_35     =   "frmMain.frx":1D67F
      submenu_item_count_35=   2
      submenu_item_count_35=   2
      submenuid_35_0  =   0
      submenutext_35_1=   "���ɴ����ļ� (&C)"
      submenuid_35_1  =   36
      submenutext_35_2=   "���ɿ�ִ���ļ� (&E) Ctrl+F5"
      submenuid_35_2  =   37
      menuid_36       =   35
      menutext_36     =   "���ɴ����ļ� (&C)"
      menuvisible_36  =   -1  'True
      menuicon_36     =   "frmMain.frx":1D69F
      submenuid_36_0  =   0
      menuid_37       =   36
      menutext_37     =   "���ɿ�ִ���ļ� (&E) Ctrl+F5"
      menuvisible_37  =   -1  'True
      menuicon_37     =   "frmMain.frx":1D8AF
      submenuid_37_0  =   0
      menuid_38       =   37
      menutext_38     =   "����"
      menuvisible_38  =   -1  'True
      menuicon_38     =   "frmMain.frx":1D8CF
      submenu_item_count_38=   9
      submenu_item_count_38=   9
      submenuid_38_0  =   0
      submenutext_38_1=   "����"
      submenuid_38_1  =   39
      submenutext_38_2=   "���� (&R)           F5"
      submenuid_38_2  =   53
      submenutext_38_3=   "�ж� (&B)           Ctrl+Alt+Break"
      submenuid_38_3  =   54
      submenutext_38_4=   "ֹͣ (&E)           Shift+F5"
      submenuid_38_4  =   55
      submenutext_38_5=   "�������� (&S)       Ctrl+Shift+F5"
      submenuid_38_5  =   56
      submenutext_38_6=   "-"
      submenuid_38_6  =   57
      submenutext_38_7=   "�����ִ��         F11"
      submenuid_38_7  =   58
      submenutext_38_8=   "�����ִ��         F10"
      submenuid_38_8  =   59
      submenutext_38_9=   "ִ�е�����         Shift+F11"
      submenuid_38_9  =   60
      menuid_39       =   38
      menutext_39     =   "����"
      menuvisible_39  =   -1  'True
      menuicon_39     =   "frmMain.frx":1D8EF
      submenu_item_count_39=   13
      submenu_item_count_39=   13
      submenuid_39_0  =   0
      submenutext_39_1=   "�ϵ��б� (&B)       Ctrl+Alt+B"
      submenuid_39_1  =   40
      submenutext_39_2=   "-"
      submenuid_39_2  =   41
      submenutext_39_3=   "���Ӵ��� (&W)       Ctrl+Alt+W"
      submenuid_39_3  =   42
      submenutext_39_4=   "���� (&L)           Ctrl+Alt+L"
      submenuid_39_4  =   43
      submenutext_39_5=   "�������� (&I)       Ctrl+Alt+I"
      submenuid_39_5  =   44
      submenutext_39_6=   "-"
      submenuid_39_6  =   45
      submenutext_39_7=   "���ö�ջ (&C)       Ctrl+Alt+C"
      submenuid_39_7  =   46
      submenutext_39_8=   "�߳� (&T)           Ctrl+Alt+T"
      submenuid_39_8  =   47
      submenutext_39_9=   "ģ�� (&M)           Ctrl+Alt+M"
      submenuid_39_9  =   48
      submenutext_39_10=   "-"
      submenuid_39_10 =   49
      submenutext_39_11=   "�ڴ� (&E)           Ctrl+Alt+E"
      submenuid_39_11 =   50
      submenutext_39_12=   "�Ĵ��� (&R)         Ctrl+Alt+R"
      submenuid_39_12 =   51
      submenutext_39_13=   "����� (&D)         Ctrl+Alt+D"
      submenuid_39_13 =   52
      menuid_40       =   39
      menutext_40     =   "�ϵ��б� (&B)       Ctrl+Alt+B"
      menucheckbox_40 =   -1  'True
      menuvisible_40  =   -1  'True
      menuicon_40     =   "frmMain.frx":1D90F
      submenuid_40_0  =   0
      menuid_41       =   40
      menutext_41     =   "-"
      menuvisible_41  =   -1  'True
      menuicon_41     =   "frmMain.frx":1DA16
      submenuid_41_0  =   0
      menuid_42       =   41
      menutext_42     =   "���Ӵ��� (&W)       Ctrl+Alt+W"
      menucheckbox_42 =   -1  'True
      menuvisible_42  =   -1  'True
      menuicon_42     =   "frmMain.frx":1DA36
      submenuid_42_0  =   0
      menuid_43       =   42
      menutext_43     =   "���� (&L)           Ctrl+Alt+L"
      menucheckbox_43 =   -1  'True
      menuvisible_43  =   -1  'True
      menuicon_43     =   "frmMain.frx":1DB6D
      submenuid_43_0  =   0
      menuid_44       =   43
      menutext_44     =   "�������� (&I)       Ctrl+Alt+I"
      menucheckbox_44 =   -1  'True
      menuvisible_44  =   -1  'True
      menuicon_44     =   "frmMain.frx":1DB8D
      submenuid_44_0  =   0
      menuid_45       =   44
      menutext_45     =   "-"
      menuvisible_45  =   -1  'True
      menuicon_45     =   "frmMain.frx":1DBAD
      submenuid_45_0  =   0
      menuid_46       =   45
      menutext_46     =   "���ö�ջ (&C)       Ctrl+Alt+C"
      menucheckbox_46 =   -1  'True
      menuvisible_46  =   -1  'True
      menuicon_46     =   "frmMain.frx":1DBCD
      submenuid_46_0  =   0
      menuid_47       =   46
      menutext_47     =   "�߳� (&T)           Ctrl+Alt+T"
      menucheckbox_47 =   -1  'True
      menuvisible_47  =   -1  'True
      menuicon_47     =   "frmMain.frx":1DBED
      submenuid_47_0  =   0
      menuid_48       =   47
      menutext_48     =   "ģ�� (&M)           Ctrl+Alt+M"
      menucheckbox_48 =   -1  'True
      menuvisible_48  =   -1  'True
      menuicon_48     =   "frmMain.frx":1DE66
      submenuid_48_0  =   0
      menuid_49       =   48
      menutext_49     =   "-"
      menuvisible_49  =   -1  'True
      menuicon_49     =   "frmMain.frx":1DF2A
      submenuid_49_0  =   0
      menuid_50       =   49
      menutext_50     =   "�ڴ� (&E)           Ctrl+Alt+E"
      menucheckbox_50 =   -1  'True
      menuvisible_50  =   -1  'True
      menuicon_50     =   "frmMain.frx":1DF4A
      submenuid_50_0  =   0
      menuid_51       =   50
      menutext_51     =   "�Ĵ��� (&R)         Ctrl+Alt+R"
      menucheckbox_51 =   -1  'True
      menuvisible_51  =   -1  'True
      menuicon_51     =   "frmMain.frx":1E02D
      submenuid_51_0  =   0
      menuid_52       =   51
      menutext_52     =   "����� (&D)         Ctrl+Alt+D"
      menucheckbox_52 =   -1  'True
      menuvisible_52  =   -1  'True
      menuicon_52     =   "frmMain.frx":1E04D
      submenuid_52_0  =   0
      menuid_53       =   52
      menutext_53     =   "���� (&R)           F5"
      menuvisible_53  =   -1  'True
      menuicon_53     =   "frmMain.frx":1E103
      submenuid_53_0  =   0
      menuid_54       =   53
      menutext_54     =   "�ж� (&B)           Ctrl+Alt+Break"
      menuvisible_54  =   -1  'True
      menuicon_54     =   "frmMain.frx":1E388
      submenuid_54_0  =   0
      menuid_55       =   54
      menutext_55     =   "ֹͣ (&E)           Shift+F5"
      menuvisible_55  =   -1  'True
      menuicon_55     =   "frmMain.frx":1E43F
      submenuid_55_0  =   0
      menuid_56       =   55
      menutext_56     =   "�������� (&S)       Ctrl+Shift+F5"
      menuvisible_56  =   -1  'True
      menuicon_56     =   "frmMain.frx":1E517
      submenuid_56_0  =   0
      menuid_57       =   56
      menutext_57     =   "-"
      menuvisible_57  =   -1  'True
      menuicon_57     =   "frmMain.frx":1E537
      submenuid_57_0  =   0
      menuid_58       =   57
      menutext_58     =   "�����ִ��         F11"
      menuvisible_58  =   -1  'True
      menuicon_58     =   "frmMain.frx":1E557
      submenuid_58_0  =   0
      menuid_59       =   58
      menutext_59     =   "�����ִ��         F10"
      menuvisible_59  =   -1  'True
      menuicon_59     =   "frmMain.frx":1E577
      submenuid_59_0  =   0
      menuid_60       =   59
      menutext_60     =   "ִ�е�����         Shift+F11"
      menuvisible_60  =   -1  'True
      menuicon_60     =   "frmMain.frx":1E597
      submenuid_60_0  =   0
      menuid_61       =   60
      menutext_61     =   "����"
      menuvisible_61  =   -1  'True
      menuicon_61     =   "frmMain.frx":1E5B7
      submenu_item_count_61=   5
      submenu_item_count_61=   5
      submenuid_61_0  =   0
      submenutext_61_1=   "���ڹ��� (&W)"
      submenuid_61_1  =   62
      submenutext_61_2=   "��Ϣ���� (&M)"
      submenuid_61_2  =   63
      submenutext_61_3=   "���� (&P)"
      submenuid_61_3  =   64
      submenutext_61_4=   "-"
      submenuid_61_4  =   65
      submenutext_61_5=   "���� (&O)"
      submenuid_61_5  =   66
      menuid_62       =   61
      menutext_62     =   "���ڹ��� (&W)"
      menuvisible_62  =   -1  'True
      menuicon_62     =   "frmMain.frx":1E5D7
      submenuid_62_0  =   0
      menuid_63       =   62
      menutext_63     =   "��Ϣ���� (&M)"
      menuvisible_63  =   -1  'True
      menuicon_63     =   "frmMain.frx":1E6BB
      submenuid_63_0  =   0
      menuid_64       =   63
      menutext_64     =   "���� (&P)"
      menuvisible_64  =   -1  'True
      menuicon_64     =   "frmMain.frx":1E78F
      submenuid_64_0  =   0
      menuid_65       =   64
      menutext_65     =   "-"
      menuvisible_65  =   -1  'True
      menuicon_65     =   "frmMain.frx":1EA08
      submenuid_65_0  =   0
      menuid_66       =   65
      menutext_66     =   "���� (&O)"
      menuvisible_66  =   -1  'True
      menuicon_66     =   "frmMain.frx":1EA28
      submenuid_66_0  =   0
      menuid_67       =   66
      menutext_67     =   "����"
      menuvisible_67  =   -1  'True
      menuicon_67     =   "frmMain.frx":1EC29
      submenu_item_count_67=   3
      submenu_item_count_67=   3
      submenuid_67_0  =   0
      submenutext_67_1=   "�����ĵ� (&D)       F1"
      submenuid_67_1  =   68
      submenutext_67_2=   "ʾ������ (&E)"
      submenuid_67_2  =   69
      submenutext_67_3=   "�����Ͽؼ��� (&A) Ctrl+F1"
      submenuid_67_3  =   70
      menuid_68       =   67
      menutext_68     =   "�����ĵ� (&D)       F1"
      menuvisible_68  =   -1  'True
      menuicon_68     =   "frmMain.frx":1EC49
      submenuid_68_0  =   0
      menuid_69       =   68
      menutext_69     =   "ʾ������ (&E)"
      menuvisible_69  =   -1  'True
      menuicon_69     =   "frmMain.frx":1ED39
      submenuid_69_0  =   0
      menuid_70       =   69
      menutext_70     =   "�����Ͽؼ��� (&A) Ctrl+F1"
      menuvisible_70  =   -1  'True
      menuicon_70     =   "frmMain.frx":1EF14
      submenuid_70_0  =   0
      levels_count    =   70
      levels_2        =   1
      levels_3        =   1
      levels_4        =   1
      levels_5        =   1
      levels_6        =   1
      levels_7        =   1
      levels_9        =   1
      levels_10       =   1
      levels_11       =   1
      levels_12       =   1
      levels_13       =   1
      levels_14       =   1
      levels_15       =   1
      levels_16       =   1
      levels_17       =   1
      levels_18       =   1
      levels_19       =   1
      levels_20       =   1
      levels_21       =   1
      levels_22       =   1
      levels_23       =   1
      levels_24       =   1
      levels_25       =   1
      levels_26       =   1
      levels_27       =   1
      levels_29       =   1
      levels_30       =   1
      levels_31       =   1
      levels_32       =   1
      levels_33       =   1
      levels_34       =   1
      levels_36       =   1
      levels_37       =   1
      levels_39       =   1
      levels_40       =   2
      levels_41       =   2
      levels_42       =   2
      levels_43       =   2
      levels_44       =   2
      levels_45       =   2
      levels_46       =   2
      levels_47       =   2
      levels_48       =   2
      levels_49       =   2
      levels_50       =   2
      levels_51       =   2
      levels_52       =   2
      levels_53       =   1
      levels_54       =   1
      levels_55       =   1
      levels_56       =   1
      levels_57       =   1
      levels_58       =   1
      levels_59       =   1
      levels_60       =   1
      levels_62       =   1
      levels_63       =   1
      levels_64       =   1
      levels_65       =   1
      levels_66       =   1
      levels_68       =   1
      levels_69       =   1
      levels_70       =   1
   End
   Begin VB.PictureBox picToolBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00302D2D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   372
      ScaleWidth      =   16848
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
         Size            =   7.8
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
      ScaleWidth      =   16848
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
         ScaleHeight     =   2052
         ScaleWidth      =   5652
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
            _extentx        =   14415
            _extenty        =   6371
         End
      End
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorderSizer 
      Left            =   16200
      Top             =   7200
      _extentx        =   677
      _extenty        =   677
      thickness       =   3
      minwidth        =   400
      minheight       =   100
      transparency    =   1
      usesetparent    =   0   'False
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   15600
      Top             =   7200
      _extentx        =   677
      _extenty        =   677
      minwidth        =   400
      minheight       =   100
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16845
      _extentx        =   29718
      _extenty        =   868
      font            =   "frmMain.frx":1F632
      caption         =   "�Ͽؼ���"
      maxbuttonvisible=   0   'False
      minbuttonvisible=   0   'False
      bindcaption     =   -1  'True
      picture         =   "frmMain.frx":1F666
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   14160
      Top             =   7320
      _Version        =   983043
      _ExtentX        =   508
      _ExtentY        =   508
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DockingPane 
      Left            =   14880
      Top             =   7320
      _Version        =   983043
      _ExtentX        =   508
      _ExtentY        =   508
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
Dim NewCreateWindow         As frmCreateOptions                                                '���½���Ŀ������

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
    
    '���ò˵���
    'ToDo
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
    NoSkinMsgBox ShowOpen(Me.hWnd, "Dilidi - Open", "ϴƨƨ�ļ�(*.cpp)" & vbNullChar & "*.cpp")
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
    SetWindowPos frmSaveBox.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE    '�ñ���Ի����ö�
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
    NoSkinMsgBox ShowSave(Me.hWnd, "Shar.cpp", "Save", "fsaf(*.cpp)" & vbNullChar & "*.cpp")
End Sub

'����:      ���½���Ŀ���˵�
Private Sub mnuNewProject_Click()
    'frmCreate.Hide                                                                      '��������������ġ��½���Ŀ������
    If Not NewCreateWindow Is Nothing Then                                              'ж�ص���һ�����½���Ŀ������
        Unload NewCreateWindow
        Set NewCreateWindow = Nothing
    End If
    Set NewCreateWindow = New frmCreateOptions
    Me.Enabled = False
    Me.DarkWindowBorderSizer.Bind = False
    SetParent NewCreateWindow.hWnd, 0
    NewCreateWindow.TypeOption(2).Focused = True
    NewCreateWindow.RefreshName
    NewCreateWindow.Move Screen.Width / 2 - frmCreate.Width / 2, Screen.Height / 2 - frmCreate.Height / 2
    NewCreateWindow.DarkTitleBar_NoDrop.Visible = True
    NewCreateWindow.DarkWindowBorder.Bind = True
    NewCreateWindow.Show
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
    
    If CurrState = 2 Then                                                       '�����ж�״̬
        Call ClearDebugWindows                                                      '������е��Դ��ڵ���Ϣ
        GdbPipe.DosInput "continue" & vbCrLf                                        '���ͼ�����������
        Exit Sub
    End If
    
    If IsSaveRequired() Then                                                    '��ʾ�����ļ�
        MsgBoxRtn = NoSkinMsgBox(Lang_Main_SaveBeforeCompile, vbQuestion Or vbYesNoCancel, Lang_Msgbox_Confirm)
        If MsgBoxRtn = vbYes Then
            SaveRtn = mnuSave_Click()
            If SaveRtn = 2 Then                                                         '����ʱ����
                If NoSkinMsgBox(Lang_Main_SaveFailedBeforeCompile, vbQuestion Or vbYesNo, Lang_Msgbox_Confirm) = vbNo Then
                    frmOutput.OutputLog Lang_Main_DebugAborted
                    Exit Sub
                End If
            ElseIf SaveRtn = 3 Or SaveRtn = 4 Then                                      '�û�ȡ������ ���� �û�ѡ�񲻱��� ��ȡ���������Ĳ���
                frmOutput.OutputLog Lang_Main_DebugAborted
                Exit Sub
            End If
        ElseIf MsgBoxRtn = vbCancel Then
            frmOutput.OutputLog Lang_Main_DebugAborted                              '�û�ѡ��ȡ������
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
            Exit Sub
        End If
    End If
    frmOutput.edOutput.Text = ""                                                '������
    '======================================================================
    
    'ʹ��g++���б���
    '                   ��ת����ǰ�������ڵ��̷�                    ������g++.exe���б���      ������Ϊ���Գ���           �����е�cpp�����ļ�
    '�����ʽ: cmd /c ���̷���: && cd "��g++.exe����Ŀ¼��" && "��g++.exe·����" [-mwindows] -g -o "�����·����" "��cpp�ļ�1��" "��cpp�ļ�2��"
    '                                       ��ת��g++.exe���ڵ�Ŀ¼                 ���Ƿ�Ϊ�����г���   �������EXE���·��
    GccCmdLine = "cmd /c " & Left(GccPath, 1) & ": && " & _
        "cd """ & Left(GccPath, InStrRev(GccPath, "\")) & """ && " & _
        """" & GccPath & """ -static -g -o """ & ExePath & """"
    For i = 0 To UBound(CurrentProject.Files)
        If Not CurrentProject.Files(i).IsHeaderFile Then
            GccCmdLine = GccCmdLine & " """ & CurrentProject.Files(i).FilePath & """"
        End If
    Next i
    If GccPipe.InitDosIO(GccCmdLine) = 0 Then                                   'g++�ܵ�����ʧ��
        frmOutput.OutputLog Lang_Main_GccStartFailed & GccCmdLine
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
            If GccOutputContent(i) <> "" Then                                   '������ǿ���
                frmOutput.OutputLog GccOutputContent(i)
            End If
        Next i
    End If
    If Dir(ExePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem) = "" Then   '���exe·�������ڣ���˵�����벻�ɹ�
        frmOutput.OutputLog Lang_Main_CompileFailed
        Exit Sub
    Else
        frmOutput.OutputLog Lang_Main_CompileSucceed & ExePath
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
        Exit Sub
    End If
    frmOutput.OutputLog Lang_Main_GdbSucceed & GdbPipe.dwProcessId & "(" & Hex(GdbPipe.dwProcessId) & ")"
    '======================================================================
    
    frmOutput.OutputLog Lang_Main_GdbLoadingSymbols_1 & ExePath & Lang_Main_GdbLoadingSymbols_2
    GdbPipe.DosInput "file """ & Replace(ExePath, "\", "/") & """" & vbCrLf     '��exe�ļ���ȡ����
    GdbPipe.DosOutput PipeOutput, "(gdb) ", 5000                                '��ȡgdb�����
    If InStr(PipeOutput, "no debugging symbols found") <> 0 Or _
        InStr(PipeOutput, "No such file or directory") <> 0 Then                    'gdb�����no debugging symbols found�����ߡ�No such file or directory�������ط���ʧ��
        frmOutput.OutputLog CStr(Split(PipeOutput, vbCrLf)(0))                      '������ط��ŵĴ���
        If NoSkinMsgBox(Lang_Main_GdbLoadSymbolsFailure_1 & ExePath & Lang_Main_GdbLoadSymbolsFailure_2, vbExclamation Or vbYesNo, Lang_Msgbox_Confirm) = vbNo Then
            TerminateProcess DebugProgramInfo.hProcess, 0                               'ɱ�������Խ��̣���������
            Set GdbPipe = Nothing                                                       '�ر�gdb�ܵ�
            frmOutput.OutputLog Lang_Main_DebugAborted
            Exit Sub
        End If
    End If
    '======================================================================
    
    GdbPipe.DosInput "set pagination off" & vbCrLf                              '�ر�gdb��"Type to continue, or q to quit"��Ϣ
    GdbPipe.DosInput "set print repeats 0" & vbCrLf                             '�ر�gdb�����ظ�������Ԫ�صġ�<repeats n times>�����
    '======================================================================
    
    frmOutput.OutputLog Lang_Main_GdbAttaching
    GdbPipe.DosInput "attach " & DebugProgramInfo.dwProcessId & vbCrLf          '���ӵ������Խ���
    GdbPipe.DosOutput PipeOutput, "(gdb) "                                      '��ȡgdb�����
    If InStr(PipeOutput, "Can't attach") <> 0 Then                              'gdb�����Can't attach to process.�������ӽ���ʧ��
        TerminateProcess DebugProgramInfo.hProcess, 0                               'ɱ�������Խ��̣���������
        Set GdbPipe = Nothing                                                       '�ر�gdb�ܵ�
        frmOutput.OutputLog Lang_Main_GdbAttachFailed_1 & DebugProgramInfo.dwProcessId & "(" & Hex(DebugProgramInfo.dwProcessId) & ") " & Lang_Main_GdbAttachFailed_2
        frmOutput.OutputLog Lang_Main_DebugAborted
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
                    frmOutput.OutputLog Lang_Main_GdbBreakpointError_1 & CurrentProject.Files(i).FilePath & _
                        Lang_Main_GdbBreakpointError_2 & Replace(Split(PipeOutput, " in file """)(0), "No line ", "") & Lang_Main_GdbBreakpointError_3
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
    SetParent frmCreate.hWnd, Me.picClientArea.hWnd                                     '�á��½���Ŀ����Ϊ��������Ӵ���
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
    
    '����LOGO
    frmStartupLogo.Show
    SetWindowPos frmStartupLogo.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
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
    
    Me.DockingPane.AttachToWindow Me.picClientArea.hWnd                                                                 '�󶨹�����
    ClientHeight = Me.picClientArea.ScaleHeight / Screen.TwipsPerPixelY
    ClientWidth = Me.picClientArea.ScaleWidth / Screen.TwipsPerPixelX
    Me.DockingPane.CreatePane 1, 100, ClientHeight, DockLeftOf                                                          '�ؼ���
    Me.DockingPane(1).Handle = frmControlBox.hWnd
    Me.DockingPane(1).Title = Lang_ControlBox_Caption
    Me.DockingPane.CreatePane 2, 250, ClientHeight / 2, DockRightOf                                                     '����
    Me.DockingPane(2).Handle = frmProperties.hWnd
    Me.DockingPane(2).Title = Lang_Properties_Caption
    Me.DockingPane.CreatePane 3, 250, ClientHeight / 2, DockRightOf                                                     '������Դ������
    Me.DockingPane(3).Handle = frmSolutionExplorer.hWnd
    Me.DockingPane(3).Title = Lang_SolutionExplorer_Caption
    Me.DockingPane.CreatePane 4, ClientWidth / 2, 120, DockBottomOf Or DockLeftOf                                       '�����б�
    Me.DockingPane(4).Handle = frmErrorList.hWnd
    Me.DockingPane(4).Title = Lang_ErrorList_Caption
    Me.DockingPane.CreatePane 5, ClientWidth / 2, 120, DockBottomOf Or DockRightOf                                      '���
    Me.DockingPane(5).Handle = frmOutput.hWnd
    Me.DockingPane(5).Title = Lang_Output_Caption
    Me.DockingPane.CreatePane 6, ClientWidth / 2, 120, DockBottomOf Or DockRightOf                                      '�ϵ��б�
    Me.DockingPane(6).Handle = frmBreakpoints.hWnd
    Me.DockingPane(6).Title = Lang_Breakpoints_Caption
    Me.DockingPane.CreatePane 7, ClientWidth / 2, 120, DockBottomOf                                                     '���Ӵ���
    Me.DockingPane(7).Handle = frmWatch.hWnd
    Me.DockingPane(7).Title = Lang_Watch_Caption
    Me.DockingPane.CreatePane 8, ClientWidth / 2, 120, DockBottomOf                                                     '����
    Me.DockingPane(8).Handle = frmLocals.hWnd
    Me.DockingPane(8).Title = Lang_Locals_Caption
    Me.DockingPane.CreatePane 9, ClientWidth / 2, 120, DockBottomOf                                                     '��������
    Me.DockingPane(9).Handle = frmImmediate.hWnd
    Me.DockingPane(9).Title = Lang_Immediate_Caption
    Me.DockingPane.CreatePane 10, ClientWidth / 2, 120, DockBottomOf                                                    '���ö�ջ
    Me.DockingPane(10).Handle = frmCallStack.hWnd
    Me.DockingPane(10).Title = Lang_CallStack_Caption
    Me.DockingPane.CreatePane 11, ClientWidth / 2, 120, DockBottomOf                                                    '�߳�
    Me.DockingPane(11).Handle = frmThreads.hWnd
    Me.DockingPane(11).Title = Lang_Threads_Caption
    Me.DockingPane.CreatePane 12, ClientWidth / 2, 120, DockBottomOf                                                    'ģ��
    Me.DockingPane(12).Handle = frmModules.hWnd
    Me.DockingPane(12).Title = Lang_Modules_Caption
    Me.DockingPane.CreatePane 13, ClientWidth / 2, 250, DockBottomOf                                                    '�ڴ�
    Me.DockingPane(13).Handle = frmMemory.hWnd
    Me.DockingPane(13).Title = Lang_Memory_Caption
    Me.DockingPane.CreatePane 14, ClientWidth / 2, 250, DockBottomOf                                                    '�Ĵ���
    Me.DockingPane(14).Handle = frmRegisters.hWnd
    Me.DockingPane(14).Title = Lang_Registers_Caption
    Me.DockingPane.CreatePane 15, ClientWidth / 2, 250, DockBottomOf                                                    '�����
    Me.DockingPane(15).Handle = frmDisassembly.hWnd
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
    '    MsgBox "����Ƥ��ʧ�ܣ�", vbCritical, Lang_Msgbox_Error todo: multi language
    'End If
    
    'todo ɾ��-----------------
    GccPath = "C:\Program Files\mingw-w64\x86_64-8.1.0-win32-seh-rt_v6-rev0\mingw64\bin\g++.exe"                                                                '����g++·��
    GdbPath = "C:\Program Files\mingw-w64\x86_64-8.1.0-win32-seh-rt_v6-rev0\mingw64\bin\gdb.exe"                                                                '����gdb·��
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
    SetPropA Me.hWnd, "WindowObj", lpObj                                                                                '��¼���ڵ������ַ�������໯ж�ش�����
    'SetPropA Me.hWnd, "PrevWndProc", SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf MainWindowMaximizeCloseFixProc)    '[ToDo]
    
    '��ʾ����ҳ��
    Call ShowStartupPage
    picToolBar.Move 0, Me.DarkMenu.Top + Me.DarkMenu.Height
    Me.picClientArea.Move 0, Me.picToolBar.Top + Me.picToolBar.Height
    
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
    
    '�ָ��������໯
    SetWindowLongA Me.hWnd, GWL_WNDPROC, GetPropA(Me.hWnd, "PrevWndProc")
    
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
        If GetPropA(Target.hWnd, "Parent") = Me.picWindowClientArea.hWnd Then
            GetWindowPlacement Target.hWnd, wp
            If wp.ShowCmd = SW_MAXIMIZE Then
                ShowWindow Target.hWnd, SW_HIDE
                ShowWindow Target.hWnd, SW_MAXIMIZE
            End If
        End If
    Next Target
End Sub

Private Sub picClientArea_Click()
'    me.line
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
    Dim i                           As Long
    
    If GdbPipe.DosOutput(PipeOutput, "(gdb) ") = 1 Then                             '��ȡgdb�Ƿ�������Ϣ
        PipeOutputLine = Split(PipeOutput, vbCrLf)                                      '�ָ��gdb�����ÿһ��
        For i = 0 To UBound(PipeOutputLine)                                             '���������ÿһ��
            PipeOutput = PipeOutputLine(i)
            If PipeOutput Like "*Breakpoint *,*at*" Then                                    '�ϵ�������Ϣ��Breakpoint *, * at *:*��
                Dim BreakpointIndex     As Long                                                 '��ǰ���еĶϵ����ţ�gdb��
                Dim bSourceFileFound    As Boolean                                              '�ܷ��ҵ���Ӧ�Ĵ����ļ�
                Dim SourceLn            As Long                                                 '��Ӧ�Ĵ����к�
                
                CurrState = 2                                                                   '���µ���״̬
                SplitTmp = Split(PipeOutput, "Breakpoint ")(1)                                  '��Breakpoint [*, * at *:*]��
                BreakpointIndex = CLng(Split(SplitTmp, ", ")(0))                                '��[*], * at *:*��
                SplitTmp = Right(SplitTmp, Len(SplitTmp) - InStr(SplitTmp, " at "))             '��Breakpoint *, * at [*:*]��
                SourceLn = CLng(Right(SplitTmp, Len(SplitTmp) - InStrRev(SplitTmp, ":")))       '��*:[*]��
                
                '���û�ж�Ӧ�Ĵ��봰�ھʹ���һ���µģ��еĻ����л���ȥ���ⲿ�ִ����frmSolutionExplorer��SolutionTreeView_DoubleClick����
                '�Է���һ����ʱ���ȡ�ϵ��Ӧ��gdb�ϵ��ʱ����������GdbBreakpoints��ӳ����ȱ©�������ᵼ�´���һ����Ч�Ĵ��봰��
                If BreakpointIndex <= UBound(GdbBreakpoints) Then
                    With CurrentProject.Files(GdbBreakpoints(BreakpointIndex).FileIndex)
                        If .TargetWindow Is Nothing Then
                            Dim NewCodeWindow   As frmCodeWindow
                            Dim FileData        As String
                            Dim tmpData         As String
                            
                            Set NewCodeWindow = CreateNewCodeWindow(GdbBreakpoints(BreakpointIndex).FileIndex)  '�����µĴ��봰�岢���ð󶨵��ļ����
                            NewCodeWindow.Caption = GetFileName(.FilePath)
                            
                            Err.Clear
                            Open .FilePath For Input As #1                                                      '���Դ򿪶�Ӧ�Ĵ����ļ�
                                If Err.Number <> 0 Then
                                    Close #1
                                    NoSkinMsgBox Lang_Main_Debug_OpenSourceFailure & .FilePath, vbExclamation, Lang_Msgbox_Error
                                Else
                                    Do While Not EOF(1)
                                        Line Input #1, tmpData
                                        FileData = FileData & tmpData & vbCrLf
                                    Loop
                                End If
                            Close #1
                            
                            Me.TabBar.AddForm NewCodeWindow
                        Else
                            frmMain.TabBar.SwitchToByForm .TargetWindow                                         '�л�����Ӧ�Ĵ��봰��
                        End If
                        .TargetWindow.SyntaxEdit.CurrPos.Row = SourceLn                                     '��ת����Ӧ�Ĵ�����
                        .TargetWindow.BreakLine = SourceLn
                        .TargetWindow.SyntaxEdit.SetFocus
                        .TargetWindow.RedrawBreakpoints                                                 '�����ж��е�С��ͷ
                    End With
                End If
                frmOutput.OutputLog Lang_Main_Debug_BreakpointHit & ": " & _
                    CurrentProject.Files(GdbBreakpoints(BreakpointIndex).FileIndex).FilePath & ":" & SourceLn
                
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
            ElseIf PipeOutput Like "Program exited *" Then                                  '�����˳���Ϣ��Program exited *.�����ɰ�gdb��
                SplitTmp = Right(PipeOutput, Len(PipeOutput) - InStrRev(PipeOutput, " "))       '��Program exited with code [*.]����Program exited [normally.]��
                SplitTmp = Left(SplitTmp, Len(SplitTmp) - 1)                                    '��[*].����[normally].��
                
                If SplitTmp = "normally" Then                                                   '������������������0
                    ExitCode = 0
                Else                                                                            '����Ͱѷ���ֵ���˽��ƣ�ת��ʮ����
                    ExitCode = CLng("&O" & SplitTmp)
                End If
                Call ProcessExitedHandler(ExitCode)
            End If
        Next i
    Else
        Me.tmrCheckProcess.Enabled = False
        MsgBox "gdb BOOMED!"
    End If
End Sub

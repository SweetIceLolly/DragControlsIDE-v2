VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.DockingPane.v15.3.1.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "拖控件大法"
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
      _extentx        =   29713
      _extenty        =   609
      font            =   "frmMain.frx":1BCC2
      menu_item_count =   70
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
      menuid_1        =   0
      menutext_1      =   "文件"
      menuvisible_1   =   -1  'True
      menuicon_1      =   "frmMain.frx":1BCF6
      submenu_item_count_1=   6
      submenuid_1_0   =   0
      submenutext_1_1 =   "新建项目 (&N)       Ctrl+N"
      submenuid_1_1   =   2
      submenutext_1_2 =   "加载项目 (&O)       Ctrl+O"
      submenuid_1_2   =   3
      submenutext_1_3 =   "保存 (&S)           Ctrl+S"
      submenuid_1_3   =   4
      submenutext_1_4 =   "另存为 (&A)         Ctrl+Shift+S"
      submenuid_1_4   =   5
      submenutext_1_5 =   "-"
      submenuid_1_5   =   6
      submenutext_1_6 =   "退出 (&E)"
      submenuid_1_6   =   7
      menuid_2        =   1
      menutext_2      =   "新建项目 (&N)       Ctrl+N"
      menuvisible_2   =   -1  'True
      menuicon_2      =   "frmMain.frx":1BD16
      submenuid_2_0   =   0
      menuid_3        =   2
      menutext_3      =   "加载项目 (&O)       Ctrl+O"
      menuvisible_3   =   -1  'True
      menuicon_3      =   "frmMain.frx":1BD36
      submenuid_3_0   =   0
      menuid_4        =   3
      menutext_4      =   "保存 (&S)           Ctrl+S"
      menuvisible_4   =   -1  'True
      menuicon_4      =   "frmMain.frx":1BD56
      submenuid_4_0   =   0
      menuid_5        =   4
      menutext_5      =   "另存为 (&A)         Ctrl+Shift+S"
      menuvisible_5   =   -1  'True
      menuicon_5      =   "frmMain.frx":1BD76
      submenuid_5_0   =   0
      menuid_6        =   5
      menutext_6      =   "-"
      menuvisible_6   =   -1  'True
      menuicon_6      =   "frmMain.frx":1BD96
      submenuid_6_0   =   0
      menuid_7        =   6
      menutext_7      =   "退出 (&E)"
      menuvisible_7   =   -1  'True
      menuicon_7      =   "frmMain.frx":1BDB6
      submenuid_7_0   =   0
      menuid_8        =   7
      menutext_8      =   "编辑"
      menuvisible_8   =   -1  'True
      menuicon_8      =   "frmMain.frx":1BDD6
      submenu_item_count_8=   19
      submenuid_8_0   =   0
      submenutext_8_1 =   "撤销 (&U)           Ctrl+Z"
      submenuid_8_1   =   9
      submenutext_8_2 =   "重复 (&R)           Ctrl+Y"
      submenuid_8_2   =   10
      submenutext_8_3 =   "-"
      submenuid_8_3   =   11
      submenutext_8_4 =   "剪切 (&U)           Ctrl+X"
      submenuid_8_4   =   12
      submenutext_8_5 =   "复制 (&C)           Ctrl+C"
      submenuid_8_5   =   13
      submenutext_8_6 =   "粘贴 (&P)           Ctrl+V"
      submenuid_8_6   =   14
      submenutext_8_7 =   "全选 (&S)           Ctrl+A"
      submenuid_8_7   =   15
      submenutext_8_8 =   "删除行 (&D)         Ctrl+L"
      submenuid_8_8   =   16
      submenutext_8_9 =   "-"
      submenuid_8_9   =   17
      submenutext_8_10=   "查找 (&F)           Ctrl+F"
      submenuid_8_10  =   18
      submenutext_8_11=   "替换 (&E)           Ctrl+H"
      submenuid_8_11  =   19
      submenutext_8_12=   "-"
      submenuid_8_12  =   20
      submenutext_8_13=   "向外缩进 (&I)       Tab"
      submenuid_8_13  =   21
      submenutext_8_14=   "向内缩进 (&O)       Shift+Tab"
      submenuid_8_14  =   22
      submenutext_8_15=   "-"
      submenuid_8_15  =   23
      submenutext_8_16=   "添加/移除断点 (&B)  F9"
      submenuid_8_16  =   24
      submenutext_8_17=   "清除所有断点 (&M)"
      submenuid_8_17  =   25
      submenutext_8_18=   "-"
      submenuid_8_18  =   26
      submenutext_8_19=   "跳转到行 (&J)       Ctrl+G"
      submenuid_8_19  =   27
      menuid_9        =   8
      menutext_9      =   "撤销 (&U)           Ctrl+Z"
      menuvisible_9   =   -1  'True
      menuicon_9      =   "frmMain.frx":1BDF6
      submenuid_9_0   =   0
      menuid_10       =   9
      menutext_10     =   "重复 (&R)           Ctrl+Y"
      menuvisible_10  =   -1  'True
      menuicon_10     =   "frmMain.frx":1BE16
      submenuid_10_0  =   0
      menuid_11       =   10
      menutext_11     =   "-"
      menuvisible_11  =   -1  'True
      menuicon_11     =   "frmMain.frx":1BE36
      submenuid_11_0  =   0
      menuid_12       =   11
      menutext_12     =   "剪切 (&U)           Ctrl+X"
      menuvisible_12  =   -1  'True
      menuicon_12     =   "frmMain.frx":1BE56
      submenuid_12_0  =   0
      menuid_13       =   12
      menutext_13     =   "复制 (&C)           Ctrl+C"
      menuvisible_13  =   -1  'True
      menuicon_13     =   "frmMain.frx":1BE76
      submenuid_13_0  =   0
      menuid_14       =   13
      menutext_14     =   "粘贴 (&P)           Ctrl+V"
      menuvisible_14  =   -1  'True
      menuicon_14     =   "frmMain.frx":1BE96
      submenuid_14_0  =   0
      menuid_15       =   14
      menutext_15     =   "全选 (&S)           Ctrl+A"
      menuvisible_15  =   -1  'True
      menuicon_15     =   "frmMain.frx":1BEB6
      submenuid_15_0  =   0
      menuid_16       =   15
      menutext_16     =   "删除行 (&D)         Ctrl+L"
      menuvisible_16  =   -1  'True
      menuicon_16     =   "frmMain.frx":1BED6
      submenuid_16_0  =   0
      menuid_17       =   16
      menutext_17     =   "-"
      menuvisible_17  =   -1  'True
      menuicon_17     =   "frmMain.frx":1BEF6
      submenuid_17_0  =   0
      menuid_18       =   17
      menutext_18     =   "查找 (&F)           Ctrl+F"
      menuvisible_18  =   -1  'True
      menuicon_18     =   "frmMain.frx":1BF16
      submenuid_18_0  =   0
      menuid_19       =   18
      menutext_19     =   "替换 (&E)           Ctrl+H"
      menuvisible_19  =   -1  'True
      menuicon_19     =   "frmMain.frx":1BF36
      submenuid_19_0  =   0
      menuid_20       =   19
      menutext_20     =   "-"
      menuvisible_20  =   -1  'True
      menuicon_20     =   "frmMain.frx":1BF56
      submenuid_20_0  =   0
      menuid_21       =   20
      menutext_21     =   "向外缩进 (&I)       Tab"
      menuvisible_21  =   -1  'True
      menuicon_21     =   "frmMain.frx":1BF76
      submenuid_21_0  =   0
      menuid_22       =   21
      menutext_22     =   "向内缩进 (&O)       Shift+Tab"
      menuvisible_22  =   -1  'True
      menuicon_22     =   "frmMain.frx":1BF96
      submenuid_22_0  =   0
      menuid_23       =   22
      menutext_23     =   "-"
      menuvisible_23  =   -1  'True
      menuicon_23     =   "frmMain.frx":1BFB6
      submenuid_23_0  =   0
      menuid_24       =   23
      menutext_24     =   "添加/移除断点 (&B)  F9"
      menuvisible_24  =   -1  'True
      menuicon_24     =   "frmMain.frx":1BFD6
      submenuid_24_0  =   0
      menuid_25       =   24
      menutext_25     =   "清除所有断点 (&M)"
      menuvisible_25  =   -1  'True
      menuicon_25     =   "frmMain.frx":1BFF6
      submenuid_25_0  =   0
      menuid_26       =   25
      menutext_26     =   "-"
      menuvisible_26  =   -1  'True
      menuicon_26     =   "frmMain.frx":1C016
      submenuid_26_0  =   0
      menuid_27       =   26
      menutext_27     =   "跳转到行 (&J)       Ctrl+G"
      menuvisible_27  =   -1  'True
      menuicon_27     =   "frmMain.frx":1C036
      submenuid_27_0  =   0
      menuid_28       =   27
      menutext_28     =   "视图"
      menuvisible_28  =   -1  'True
      menuicon_28     =   "frmMain.frx":1C056
      submenu_item_count_28=   6
      submenuid_28_0  =   0
      submenutext_28_1=   "工具栏 (&T)"
      submenuid_28_1  =   29
      submenutext_28_2=   "控件箱 (&C)"
      submenuid_28_2  =   30
      submenutext_28_3=   "属性 (&P)           F4"
      submenuid_28_3  =   31
      submenutext_28_4=   "工程资源管理器 (&M)"
      submenuid_28_4  =   32
      submenutext_28_5=   "错误列表 (&E)       Ctrl+E"
      submenuid_28_5  =   33
      submenutext_28_6=   "输出 (&O)           Ctrl+Alt+O"
      submenuid_28_6  =   34
      menuid_29       =   28
      menutext_29     =   "工具栏 (&T)"
      menucheckbox_29 =   -1  'True
      menuvisible_29  =   -1  'True
      menuicon_29     =   "frmMain.frx":1C076
      submenuid_29_0  =   0
      menuid_30       =   29
      menutext_30     =   "控件箱 (&C)"
      menucheckbox_30 =   -1  'True
      menuvisible_30  =   -1  'True
      menuicon_30     =   "frmMain.frx":1C096
      submenuid_30_0  =   0
      menuid_31       =   30
      menutext_31     =   "属性 (&P)           F4"
      menucheckbox_31 =   -1  'True
      menuvisible_31  =   -1  'True
      menuicon_31     =   "frmMain.frx":1C0B6
      submenuid_31_0  =   0
      menuid_32       =   31
      menutext_32     =   "工程资源管理器 (&M)"
      menucheckbox_32 =   -1  'True
      menuvisible_32  =   -1  'True
      menuicon_32     =   "frmMain.frx":1C0D6
      submenuid_32_0  =   0
      menuid_33       =   32
      menutext_33     =   "错误列表 (&E)       Ctrl+E"
      menucheckbox_33 =   -1  'True
      menuvisible_33  =   -1  'True
      menuicon_33     =   "frmMain.frx":1C0F6
      submenuid_33_0  =   0
      menuid_34       =   33
      menutext_34     =   "输出 (&O)           Ctrl+Alt+O"
      menucheckbox_34 =   -1  'True
      menuvisible_34  =   -1  'True
      menuicon_34     =   "frmMain.frx":1C116
      submenuid_34_0  =   0
      menuid_35       =   34
      menutext_35     =   "生成"
      menuvisible_35  =   -1  'True
      menuicon_35     =   "frmMain.frx":1C136
      submenu_item_count_35=   2
      submenuid_35_0  =   0
      submenutext_35_1=   "生成代码文件 (&C)"
      submenuid_35_1  =   36
      submenutext_35_2=   "生成可执行文件 (&E) Ctrl+F5"
      submenuid_35_2  =   37
      menuid_36       =   35
      menutext_36     =   "生成代码文件 (&C)"
      menuvisible_36  =   -1  'True
      menuicon_36     =   "frmMain.frx":1C156
      submenuid_36_0  =   0
      menuid_37       =   36
      menutext_37     =   "生成可执行文件 (&E) Ctrl+F5"
      menuvisible_37  =   -1  'True
      menuicon_37     =   "frmMain.frx":1C176
      submenuid_37_0  =   0
      menuid_38       =   37
      menutext_38     =   "调试"
      menuvisible_38  =   -1  'True
      menuicon_38     =   "frmMain.frx":1C196
      submenu_item_count_38=   9
      submenuid_38_0  =   0
      submenutext_38_1=   "窗口"
      submenuid_38_1  =   39
      submenutext_38_2=   "运行 (&R)           F5"
      submenuid_38_2  =   53
      submenutext_38_3=   "中断 (&B)           Ctrl+Alt+Break"
      submenuid_38_3  =   54
      submenutext_38_4=   "停止 (&E)           Shift+F5"
      submenuid_38_4  =   55
      submenutext_38_5=   "重新运行 (&S)       Ctrl+Shift+F5"
      submenuid_38_5  =   56
      submenutext_38_6=   "-"
      submenuid_38_6  =   57
      submenutext_38_7=   "逐语句执行         F11"
      submenuid_38_7  =   58
      submenutext_38_8=   "逐过程执行         F10"
      submenuid_38_8  =   59
      submenutext_38_9=   "执行到返回         Shift+F11"
      submenuid_38_9  =   60
      menuid_39       =   38
      menutext_39     =   "窗口"
      menuvisible_39  =   -1  'True
      menuicon_39     =   "frmMain.frx":1C1B6
      submenu_item_count_39=   13
      submenuid_39_0  =   0
      submenutext_39_1=   "断点列表 (&B)       Ctrl+Alt+B"
      submenuid_39_1  =   40
      submenutext_39_2=   "-"
      submenuid_39_2  =   41
      submenutext_39_3=   "监视窗口 (&W)       Ctrl+Alt+W"
      submenuid_39_3  =   42
      submenutext_39_4=   "本地 (&L)           Ctrl+Alt+L"
      submenuid_39_4  =   43
      submenutext_39_5=   "立即窗口 (&I)       Ctrl+Alt+I"
      submenuid_39_5  =   44
      submenutext_39_6=   "-"
      submenuid_39_6  =   45
      submenutext_39_7=   "调用堆栈 (&C)       Ctrl+Alt+C"
      submenuid_39_7  =   46
      submenutext_39_8=   "线程 (&T)           Ctrl+Alt+T"
      submenuid_39_8  =   47
      submenutext_39_9=   "模块 (&M)           Ctrl+Alt+M"
      submenuid_39_9  =   48
      submenutext_39_10=   "-"
      submenuid_39_10 =   49
      submenutext_39_11=   "内存 (&E)           Ctrl+Alt+E"
      submenuid_39_11 =   50
      submenutext_39_12=   "寄存器 (&R)         Ctrl+Alt+R"
      submenuid_39_12 =   51
      submenutext_39_13=   "反汇编 (&D)         Ctrl+Alt+D"
      submenuid_39_13 =   52
      menuid_40       =   39
      menutext_40     =   "断点列表 (&B)       Ctrl+Alt+B"
      menucheckbox_40 =   -1  'True
      menuvisible_40  =   -1  'True
      menuicon_40     =   "frmMain.frx":1C1D6
      submenuid_40_0  =   0
      menuid_41       =   40
      menutext_41     =   "-"
      menuvisible_41  =   -1  'True
      menuicon_41     =   "frmMain.frx":1C1F6
      submenuid_41_0  =   0
      menuid_42       =   41
      menutext_42     =   "监视窗口 (&W)       Ctrl+Alt+W"
      menucheckbox_42 =   -1  'True
      menuvisible_42  =   -1  'True
      menuicon_42     =   "frmMain.frx":1C216
      submenuid_42_0  =   0
      menuid_43       =   42
      menutext_43     =   "本地 (&L)           Ctrl+Alt+L"
      menucheckbox_43 =   -1  'True
      menuvisible_43  =   -1  'True
      menuicon_43     =   "frmMain.frx":1C236
      submenuid_43_0  =   0
      menuid_44       =   43
      menutext_44     =   "立即窗口 (&I)       Ctrl+Alt+I"
      menucheckbox_44 =   -1  'True
      menuvisible_44  =   -1  'True
      menuicon_44     =   "frmMain.frx":1C256
      submenuid_44_0  =   0
      menuid_45       =   44
      menutext_45     =   "-"
      menuvisible_45  =   -1  'True
      menuicon_45     =   "frmMain.frx":1C276
      submenuid_45_0  =   0
      menuid_46       =   45
      menutext_46     =   "调用堆栈 (&C)       Ctrl+Alt+C"
      menucheckbox_46 =   -1  'True
      menuvisible_46  =   -1  'True
      menuicon_46     =   "frmMain.frx":1C296
      submenuid_46_0  =   0
      menuid_47       =   46
      menutext_47     =   "线程 (&T)           Ctrl+Alt+T"
      menucheckbox_47 =   -1  'True
      menuvisible_47  =   -1  'True
      menuicon_47     =   "frmMain.frx":1C2B6
      submenuid_47_0  =   0
      menuid_48       =   47
      menutext_48     =   "模块 (&M)           Ctrl+Alt+M"
      menucheckbox_48 =   -1  'True
      menuvisible_48  =   -1  'True
      menuicon_48     =   "frmMain.frx":1C2D6
      submenuid_48_0  =   0
      menuid_49       =   48
      menutext_49     =   "-"
      menuvisible_49  =   -1  'True
      menuicon_49     =   "frmMain.frx":1C2F6
      submenuid_49_0  =   0
      menuid_50       =   49
      menutext_50     =   "内存 (&E)           Ctrl+Alt+E"
      menucheckbox_50 =   -1  'True
      menuvisible_50  =   -1  'True
      menuicon_50     =   "frmMain.frx":1C316
      submenuid_50_0  =   0
      menuid_51       =   50
      menutext_51     =   "寄存器 (&R)         Ctrl+Alt+R"
      menucheckbox_51 =   -1  'True
      menuvisible_51  =   -1  'True
      menuicon_51     =   "frmMain.frx":1C336
      submenuid_51_0  =   0
      menuid_52       =   51
      menutext_52     =   "反汇编 (&D)         Ctrl+Alt+D"
      menucheckbox_52 =   -1  'True
      menuvisible_52  =   -1  'True
      menuicon_52     =   "frmMain.frx":1C356
      submenuid_52_0  =   0
      menuid_53       =   52
      menutext_53     =   "运行 (&R)           F5"
      menuvisible_53  =   -1  'True
      menuicon_53     =   "frmMain.frx":1C376
      submenuid_53_0  =   0
      menuid_54       =   53
      menutext_54     =   "中断 (&B)           Ctrl+Alt+Break"
      menuvisible_54  =   -1  'True
      menuicon_54     =   "frmMain.frx":1C396
      submenuid_54_0  =   0
      menuid_55       =   54
      menutext_55     =   "停止 (&E)           Shift+F5"
      menuvisible_55  =   -1  'True
      menuicon_55     =   "frmMain.frx":1C3B6
      submenuid_55_0  =   0
      menuid_56       =   55
      menutext_56     =   "重新运行 (&S)       Ctrl+Shift+F5"
      menuvisible_56  =   -1  'True
      menuicon_56     =   "frmMain.frx":1C3D6
      submenuid_56_0  =   0
      menuid_57       =   56
      menutext_57     =   "-"
      menuvisible_57  =   -1  'True
      menuicon_57     =   "frmMain.frx":1C3F6
      submenuid_57_0  =   0
      menuid_58       =   57
      menutext_58     =   "逐语句执行         F11"
      menuvisible_58  =   -1  'True
      menuicon_58     =   "frmMain.frx":1C416
      submenuid_58_0  =   0
      menuid_59       =   58
      menutext_59     =   "逐过程执行         F10"
      menuvisible_59  =   -1  'True
      menuicon_59     =   "frmMain.frx":1C436
      submenuid_59_0  =   0
      menuid_60       =   59
      menutext_60     =   "执行到返回         Shift+F11"
      menuvisible_60  =   -1  'True
      menuicon_60     =   "frmMain.frx":1C456
      submenuid_60_0  =   0
      menuid_61       =   60
      menutext_61     =   "工具"
      menuvisible_61  =   -1  'True
      menuicon_61     =   "frmMain.frx":1C476
      submenu_item_count_61=   5
      submenuid_61_0  =   0
      submenutext_61_1=   "窗口工具 (&W)"
      submenuid_61_1  =   62
      submenutext_61_2=   "消息拦截 (&M)"
      submenuid_61_2  =   63
      submenutext_61_3=   "进程 (&P)"
      submenuid_61_3  =   64
      submenutext_61_4=   "-"
      submenuid_61_4  =   65
      submenutext_61_5=   "设置 (&O)"
      submenuid_61_5  =   66
      menuid_62       =   61
      menutext_62     =   "窗口工具 (&W)"
      menuvisible_62  =   -1  'True
      menuicon_62     =   "frmMain.frx":1C496
      submenuid_62_0  =   0
      menuid_63       =   62
      menutext_63     =   "消息拦截 (&M)"
      menuvisible_63  =   -1  'True
      menuicon_63     =   "frmMain.frx":1C4B6
      submenuid_63_0  =   0
      menuid_64       =   63
      menutext_64     =   "进程 (&P)"
      menuvisible_64  =   -1  'True
      menuicon_64     =   "frmMain.frx":1C4D6
      submenuid_64_0  =   0
      menuid_65       =   64
      menutext_65     =   "-"
      menuvisible_65  =   -1  'True
      menuicon_65     =   "frmMain.frx":1C4F6
      submenuid_65_0  =   0
      menuid_66       =   65
      menutext_66     =   "设置 (&O)"
      menuvisible_66  =   -1  'True
      menuicon_66     =   "frmMain.frx":1C516
      submenuid_66_0  =   0
      menuid_67       =   66
      menutext_67     =   "帮助"
      menuvisible_67  =   -1  'True
      menuicon_67     =   "frmMain.frx":1C536
      submenu_item_count_67=   3
      submenuid_67_0  =   0
      submenutext_67_1=   "帮助文档 (&D)       F1"
      submenuid_67_1  =   68
      submenutext_67_2=   "示例程序 (&E)"
      submenuid_67_2  =   69
      submenutext_67_3=   "关于拖控件大法 (&A) Ctrl+F1"
      submenuid_67_3  =   70
      menuid_68       =   67
      menutext_68     =   "帮助文档 (&D)       F1"
      menuvisible_68  =   -1  'True
      menuicon_68     =   "frmMain.frx":1C556
      submenuid_68_0  =   0
      menuid_69       =   68
      menutext_69     =   "示例程序 (&E)"
      menuvisible_69  =   -1  'True
      menuicon_69     =   "frmMain.frx":1C576
      submenuid_69_0  =   0
      menuid_70       =   69
      menutext_70     =   "关于拖控件大法 (&A) Ctrl+F1"
      menuvisible_70  =   -1  'True
      menuicon_70     =   "frmMain.frx":1C596
      submenuid_70_0  =   0
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
            _extentx        =   14420
            _extenty        =   6376
         End
      End
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorderSizer 
      Left            =   16200
      Top             =   7200
      _extentx        =   847
      _extenty        =   847
      thickness       =   3
      minwidth        =   400
      minheight       =   100
      transparency    =   1
      usesetparent    =   0   'False
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   15600
      Top             =   7200
      _extentx        =   847
      _extenty        =   847
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
      _extentx        =   29713
      _extenty        =   873
      font            =   "frmMain.frx":1C5B6
      caption         =   "拖控件大法"
      maxbuttonvisible=   0   'False
      minbuttonvisible=   0   'False
      bindcaption     =   -1  'True
      picture         =   "frmMain.frx":1C5EA
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
'描述:      主窗口
'作者:      冰棍, Error 404
'文件:      frmMain.frm
'====================================================

Option Explicit

'获取窗口最大、最小化状态
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

'创建进程
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

'工程类型
'值     描述
'0      未创建工程，处于启动界面
'1      窗口程序
'2      控制台程序
'3      空白C++程序
Public ProjectType          As Integer

Public WindowObj            As Object                                                   '窗口自身
Dim NewCreateWindow         As frmCreate                                                '“新建项目”窗体

Private WithEvents GdbPipe  As clsPipe                                                  'gdb调试管道
Attribute GdbPipe.VB_VarHelpID = -1

'描述:      “加载项目”菜单
Private Sub mnuOpen_Click()
    NoSkinMsgBox ShowOpen(Me.hWnd, "Dilidi - Open", "洗屁屁文件(*.cpp)" & vbNullChar & "*.cpp")
End Sub

'描述:      “保存”菜单
Private Sub mnuSave_Click()
    On Error Resume Next
    Dim i                   As Long
    
    For i = 0 To UBound(CurrentProject.Files)                                           '检查还没有保存的文件
        If CurrentProject.Files(i).Changed = True Then                                      '逐个保存未保存的文件
            Open CurrentProject.Files(i).FilePath For Output As #1
                If Err.Number <> 0 Then                                                             '保存文件失败
                    Close #1
                    If NoSkinMsgBox(Lang_Main_SaveFailure_1 & CurrentProject.Files(i).FilePath & " :" & _
                       Err.Number & " - " & Err.Description & Lang_Main_saveFailure_2, vbExclamation, Lang_Msgbox_Error) = vbNo Then
                        Exit Sub
                    End If
                End If
                Print #1, CurrentProject.Files(i).TargetWindow.SyntaxEdit.Text
            Close #1
            CurrentProject.Files(i).Changed = False                                             '标记文件为已保存
        End If
    Next i
    
    If CurrentProject.Changed Then                                                      '如果工程文件尚未保存
        Open ProjectFilePath For Binary As #1                                               '保存工程文件
            If Err.Number <> 0 Then
                Close #1
                NoSkinMsgBox Lang_Main_SaveFailure_1 & ProjectFilePath & " :" & Err.Number & " - " & Err.Description, vbExclamation, Lang_Msgbox_Error
                Exit Sub
            End If
            'Put #1, , CurrentProject                                                           'ToDo: Make another user type or what?
        Close #1
        CurrentProject.Changed = False                                                          '标记工程文件为已保存
    End If
End Sub

'描述:      “另存为”菜单
Private Sub mnuSaveAs_Click()
    NoSkinMsgBox ShowSave(Me.hWnd, "Shar.cpp", "Save", "fsaf(*.cpp)" & vbNullChar & "*.cpp")
End Sub

'描述:      “新建项目”菜单
Private Sub mnuNewProject_Click()
    If Not NewCreateWindow Is Nothing Then                                              '卸载掉上一个“新建项目”窗体
        Unload NewCreateWindow
        Set NewCreateWindow = Nothing
    End If
    Set NewCreateWindow = New frmCreate
    Me.Enabled = False
    Me.DarkWindowBorderSizer.Bind = False
    SetParent NewCreateWindow.hWnd, 0
    NewCreateWindow.Move Screen.Width / 2 - frmCreate.Width / 2, Screen.Height / 2 - frmCreate.Height / 2
    NewCreateWindow.DarkTitleBar_NoDrop.Visible = True
    NewCreateWindow.DarkWindowBorder.Bind = True
    NewCreateWindow.Show
End Sub

'描述:      “运行”菜单
Private Sub mnuRun_Click()
    On Error Resume Next
    
    Dim GccPipe             As New clsPipe                                              'g++管道
    Dim GccCmdLine          As String                                                   'g++命令行
    Dim ExePath             As String                                                   'exe文件编译路径
    Dim PipeOutput          As String                                                   '管道输出的内容
    Dim GccOutputContent()  As String                                                   '逐行分开的g++输出内容
    Dim i                   As Long
    
    '提示保存文件
    For i = 0 To UBound(CurrentProject.Files)
        If CurrentProject.Files(i).Changed Then
            If NoSkinMsgBox(Lang_Main_SaveBeforeCompile, vbQuestion Or vbYesNo, Lang_Msgbox_Confirm) = vbYes Then
                Call mnuSave_Click
            End If
            Exit For
        End If
    Next i

    '使用g++进行编译
    '                   ↓转到当前程序所在的盘符                    ↓调用g++.exe进行编译    ↓编译为调试程序           ↓所有的cpp代码文件
    '命令格式: cmd /c 【盘符】: && cd "【g++.exe所在目录】" && "【g++.exe路径】" [-mwindows] -g -o "【输出路径】" "【cpp文件1】" "【cpp文件2】"
    '                                       ↑转到g++.exe所在的目录                 ↑是否为命令行程序   ↑编译的EXE输出路径
    frmOutput.OutputLog Lang_Main_StartingGcc
    ExePath = ProjectFolderPath & CurrentProject.ProjectName & ".exe"
    GccCmdLine = "cmd /c " & Left(GetAppPath(), 1) & ": && " & _
       "cd """ & GetAppPath() & "GCC\bin"" && " & _
       """" & GetAppPath() & "GCC\bin\g++.exe"" -g -o """ & ExePath & """"
    For i = 0 To UBound(CurrentProject.Files)
        If Not CurrentProject.Files(i).IsHeaderFile Then
            GccCmdLine = GccCmdLine & " """ & CurrentProject.Files(i).FilePath & """"
        End If
    Next i
    If GccPipe.InitDosIO(GccCmdLine) = 0 Then
        frmOutput.OutputLog Lang_Main_GccStartFailed
    End If
    frmMain.DarkMenu.HideMenu                                                           '先隐藏菜单
    Do While ProcessExists(GccPipe.hProcess)                                            '等待g++执行完成
        Sleep 50
        DoEvents
    Loop
    GccPipe.DosOutput PipeOutput, vbNullChar & vbNullChar                               '获取g++输出
    GccOutputContent = Split(PipeOutput, vbCrLf)
    If UBound(GccOutputContent) >= 0 Then
        For i = 0 To UBound(GccOutputContent)                                               '逐行输出
            If GccOutputContent(i) <> "" Then                                                   '如果不是空行
                frmOutput.OutputLog GccOutputContent(i)
            End If
        Next i
    End If
    If Dir(ExePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem) = "" Then           '如果exe路径不存在，则说明编译不成功
        frmOutput.OutputLog Lang_Main_CompileFailed
        Exit Sub
    Else
        frmOutput.OutputLog Lang_Main_CompileSucceed & ExePath
    End If
    
    '创建待调试进程。该进程启动之后会挂起，等待gdb附加
    Dim si                  As STARTUPINFO                                              '进程启动信息
    Dim sa                  As SECURITY_ATTRIBUTES                                      '安全属性
    
    With sa                                                                             '设置安全属性
        .bInheritHandle = 1
        .lpSecurityDescriptor = 0
        .nLength = Len(sa)
    End With
    If CreateProcess(0, ExePath, sa, sa, ByVal 1, _
       NORMAL_PRIORITY_CLASS Or CREATE_SUSPENDED, ByVal 0, ByVal 0, si, DebugProgramInfo) <> 1 Then
        
        frmOutput.OutputLog Lang_Main_RunFailed & ExePath & " (" & Err.LastDllError & ")"
        Exit Sub
    End If
    
    '创建gdb管道
    Set GdbPipe = New clsPipe
    If GdbPipe.InitDosIO(GetAppPath() & "GCC\gdb\gdb.exe -q -nw") = 0 Then              '创建gdb调试管道失败
        TerminateProcess DebugProgramInfo.hProcess, 0                                       '杀掉待调试进程，放弃调试
        Set GdbPipe = Nothing                                                               '关闭gdb管道
        frmOutput.OutputLog Lang_Main_GdbFailed
        Exit Sub
    End If
    GdbPipe.DosInput "attach " & DebugProgramInfo.dwProcessId & vbCrLf                  '附加到待调试进程
    GdbPipe.DosOutput PipeOutput, "(gdb) "                                              '获取gdb的输出
    If InStr(PipeOutput, "Can't attach") <> 0 Then                                      'gdb输出“Can't attach to process.”，附加进程失败
        TerminateProcess DebugProgramInfo.hProcess, 0                                       '杀掉待调试进程，放弃调试
        Set GdbPipe = Nothing                                                               '关闭gdb管道
        frmOutput.OutputLog Lang_Main_GdbAttachFailed_1 & DebugProgramInfo.dwProcessId & "(" & Hex(DebugProgramInfo.dwProcessId) & ") " & Lang_Main_GdbAttachFailed_2
        Exit Sub
    End If
    GdbPipe.DosInput "continue" & vbCrLf                                                '使目标进程继续运行
    frmOutput.OutputLog Lang_Main_DebugInfo_1 & GdbPipe.dwProcessId & "(" & Hex(GdbPipe.dwProcessId) & "); " & _
        Right(ExePath, Len(ExePath) - InStrRev(ExePath, "\")) & Lang_Main_DebugInfo_2 & DebugProgramInfo.dwProcessId & "(" & Hex(DebugProgramInfo.dwProcessId) & ")"
End Sub

'描述:      隐藏启动界面
Public Sub HideStartupPage()
    On Error Resume Next
    Unload NewCreateWindow
    Me.TabBar.Visible = True
    
    Me.DarkMenu.MenuEnabled(3) = True                                                   '保存
    Me.DarkMenu.MenuEnabled(4) = True                                                   '另存为
    Me.DarkMenu.MenuEnabled(7) = True                                                   '编辑
    Me.DarkMenu.MenuEnabled(27) = True                                                  '视图
    Me.DarkMenu.MenuEnabled(34) = True                                                  '生成
    Me.DarkMenu.MenuEnabled(37) = True                                                  '调试
End Sub

'描述:      显示启动界面
Public Sub ShowStartupPage()
    frmCreate.DarkTitleBar_NoDrop.Visible = False                                              '不显示标题栏和边框
    frmCreate.DarkWindowBorder.Bind = False
    SetParent frmCreate.hWnd, Me.picClientArea.hWnd                                     '让“新建项目”作为本窗体的子窗体
    frmCreate.Move 0, 0                                                                 '设置其位置
    frmCreate.Show
End Sub

Private Sub DarkMenu_MenuItemClicked(MenuID As Integer)
    Select Case MenuID
        Case 1                                                                          '新建
            Call mnuNewProject_Click
        
        Case 2                                                                          '加载
            Call mnuOpen_Click
        
        Case 3                                                                          '保存
            Call mnuSave_Click
        
        Case 4                                                                          '另存为
            Call mnuSaveAs_Click
        
        Case 52                                                                         '运行
            Call mnuRun_Click
        
    End Select
End Sub

Private Sub DockingPane_Resize()
    If ProjectType <> 0 Then                                                            '如果不是在启动界面的话就调整窗口活动客户区大小
        Dim cLeft   As Long, cRight   As Long, cTop   As Long, cBottom   As Long
        
        Me.DockingPane.GetClientRect cLeft, cTop, cRight, cBottom
        Me.picWindowClientArea.Move cLeft, cTop, cRight - cLeft, cBottom - cTop
        Me.TabBar.Move 0, 0, Me.picWindowClientArea.ScaleWidth, Me.picWindowClientArea.ScaleHeight
        
        Call Form_Resize                                                                    '如果窗口客户区里面有最大化的窗口，对其大小进行调整
    End If
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    
    '启动LOGO
    frmStartupLogo.Show
    SetWindowPos frmStartupLogo.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    frmStartupLogo.SetFocus

    '由于字符串资源必须比用户控件更早加载，用户控件才能使用这些字符串资源，于是放在Initialize事件而不是Load事件
    '加载字符串资源
    If Not LoadLanguage(1001) Then
        NoSkinMsgBox "加载字符串资源失败！" & Err.Number & ": " & Err.Description, vbCritical, Lang_Msgbox_Error
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    '继续隐藏窗体，等一切加载完再显示
    Me.Hide
    Me.DarkTitleBar.MinButtonVisible = True
    Me.DarkTitleBar.MaxButtonVisible = True
    Me.Caption = Lang_Application_Title
    
    '加载菜单文本
    If Not LoadLanguage(1001, True) Then
        NoSkinMsgBox "加载字符串资源失败！" & Err.Number & ": " & Err.Description, vbCritical, Lang_Msgbox_Error
    End If
    
    '调整“客户区”
    Me.picClientArea.Height = Me.ScaleHeight - Me.picClientArea.Top                                                     '主“客户区”的大小
    Me.picWindowClientArea.BackColor = Me.BackColor                                                                     '窗口客户区的颜色
    
    '创建工作区
    Dim ClientHeight        As Integer, ClientWidth             As Integer
    Dim i                   As Integer
    
    Me.DockingPane.AttachToWindow Me.picClientArea.hWnd                                                                 '绑定工作区
    ClientHeight = Me.picClientArea.ScaleHeight / Screen.TwipsPerPixelY
    ClientWidth = Me.picClientArea.ScaleWidth / Screen.TwipsPerPixelX
    Me.DockingPane.CreatePane 1, 100, ClientHeight, DockLeftOf                                                          '控件箱
    Me.DockingPane(1).Handle = frmControlBox.hWnd
    Me.DockingPane(1).Title = "控件箱"
    Me.DockingPane.CreatePane 2, 250, ClientHeight / 2, DockRightOf                                                     '属性
    Me.DockingPane(2).Handle = frmProperties.hWnd
    Me.DockingPane(2).Title = "属性"
    Me.DockingPane.CreatePane 3, 250, ClientHeight / 2, DockRightOf                                                     '工程资源管理器
    Me.DockingPane(3).Handle = frmSolutionExplorer.hWnd
    Me.DockingPane(3).Title = "工程资源管理器"
    Me.DockingPane.CreatePane 4, ClientWidth / 2, 120, DockBottomOf Or DockLeftOf                                       '错误列表
    Me.DockingPane(4).Handle = frmErrorList.hWnd
    Me.DockingPane(4).Title = "错误列表"
    Me.DockingPane.CreatePane 5, ClientWidth / 2, 120, DockBottomOf Or DockRightOf                                      '输出
    Me.DockingPane(5).Handle = frmOutput.hWnd
    Me.DockingPane(5).Title = "输出"
    Me.DockingPane.CreatePane 6, ClientWidth / 2, 120, DockBottomOf Or DockRightOf                                      '断点列表
    Me.DockingPane(6).Handle = frmBreakpoints.hWnd
    Me.DockingPane(6).Title = "断点列表"
    Me.DockingPane.CreatePane 7, ClientWidth / 2, 120, DockBottomOf                                                     '监视窗口
    Me.DockingPane(7).Handle = frmWatch.hWnd
    Me.DockingPane(7).Title = "监视窗口"
    Me.DockingPane.CreatePane 8, ClientWidth / 2, 120, DockBottomOf                                                     '本地
    Me.DockingPane(8).Handle = frmLocals.hWnd
    Me.DockingPane(8).Title = "本地"
    Me.DockingPane.CreatePane 9, ClientWidth / 2, 120, DockBottomOf                                                     '立即窗口
    Me.DockingPane(9).Handle = frmImmediate.hWnd
    Me.DockingPane(9).Title = "立即窗口"
    Me.DockingPane.CreatePane 10, ClientWidth / 2, 120, DockBottomOf                                                    '调用堆栈
    Me.DockingPane(10).Handle = frmCallStack.hWnd
    Me.DockingPane(10).Title = "调用堆栈"
    Me.DockingPane.CreatePane 11, ClientWidth / 2, 120, DockBottomOf                                                    '线程
    Me.DockingPane(11).Handle = frmThreads.hWnd
    Me.DockingPane(11).Title = "线程"
    Me.DockingPane.CreatePane 12, ClientWidth / 2, 120, DockBottomOf                                                    '模块
    Me.DockingPane(12).Handle = frmModules.hWnd
    Me.DockingPane(12).Title = "模块"
    Me.DockingPane.CreatePane 13, ClientWidth / 2, 250, DockBottomOf                                                    '内存
    Me.DockingPane(13).Handle = frmMemory.hWnd
    Me.DockingPane(13).Title = "内存"
    Me.DockingPane.CreatePane 14, ClientWidth / 2, 250, DockBottomOf                                                    '寄存器
    Me.DockingPane(14).Handle = frmRegisters.hWnd
    Me.DockingPane(14).Title = "寄存器"
    Me.DockingPane.CreatePane 15, ClientWidth / 2, 250, DockBottomOf                                                    '反汇编
    Me.DockingPane(15).Handle = frmDisassembly.hWnd
    Me.DockingPane(15).Title = "反汇编"
    For i = 1 To 15                                                                                                     '隐藏所有的Pane
        Me.DockingPane(i).Close
    Next i
    
    '设置Docking Pane的样式
    Me.DockingPane.Options.ShowDockingContextStickers = True                                                            '显示Docking stickers
    Me.DockingPane.Options.AlphaDockingContext = True                                                                   '移动的时候透明
    Me.DockingPane.Options.ThemedFloatingFrames = True                                                                  '作为弹窗时边框保持样式
    Me.DockingPane.Options.ShowContentsWhileDragging = True
    If DockingPaneGlobalSettings.ResourceImages.LoadFromFile(GetAppPath & "Skin.dll", "Office2010Black.ini") = False Then
        NoSkinMsgBox "加载样式失败！", vbCritical, "错误"
    End If
    Me.DockingPane.VisualTheme = ThemeResource                                                                          '设置为从资源文件读取样式
    Me.DockingPane.PaintManager.SplitterSize = 2                                                                        '设置分割区域的大小
    
    'If Not Me.SkinFramework.LoadSkin("Skin.cjstyles", "NormalBlue.ini") Then                                            '加载皮肤 [ToDo]
    '    NoSkinMsgBox "加载皮肤失败！", vbCritical, "错误"
    'End If
    
    '禁用不需要的菜单
    Me.DarkMenu.MenuEnabled(3) = False                                                                                  '保存
    Me.DarkMenu.MenuEnabled(4) = False                                                                                  '另存为
    Me.DarkMenu.MenuEnabled(7) = False                                                                                  '编辑
    Me.DarkMenu.MenuEnabled(27) = False                                                                                 '视图
    Me.DarkMenu.MenuEnabled(34) = False                                                                                 '生成
    Me.DarkMenu.MenuEnabled(37) = False                                                                                 '调试
    
    '设置窗口子类化，处理最大化问题及处理任务栏右键关闭
    Dim lpObj               As Long                                                                                     '指向窗口自身的物件指针
    Set WindowObj = Me
    lpObj = ObjPtr(WindowObj)                                                                                           '获取指向窗口自身的物件指针
    SetPropA Me.hWnd, "WindowObj", lpObj                                                                                '记录窗口的物件地址，供子类化卸载窗体用
    'SetPropA Me.hWnd, "PrevWndProc", SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf MainWindowMaximizeCloseFixProc)    '[ToDo]
    
    '显示启动页面
    Call ShowStartupPage
    picToolBar.Move 0, Me.DarkMenu.Top + Me.DarkMenu.Height
    Me.picClientArea.Move 0, Me.picToolBar.Top + Me.picToolBar.Height
    
    '卸载LOGO
    Unload frmStartupLogo
    Me.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    
    '恢复窗口子类化
    SetWindowLongA Me.hWnd, GWL_WNDPROC, GetPropA(Me.hWnd, "PrevWndProc")
    
    '检查“新建项目”窗口是否关闭
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
    
    '关闭所有窗口
    Dim CodeWindow  As Form
    IsExiting = True                                        '进入退出状态
    For Each CodeWindow In CodeWindows                      '卸载所有代码窗体
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
    
    '调整“客户区”的大小
    Me.picToolBar.Width = Me.ScaleWidth
    Me.picClientArea.Move 0, Me.picToolBar.Top + Me.picToolBar.Height, Me.ScaleWidth, Me.ScaleHeight - Me.picClientArea.Top
    
    '调整最大化的子窗体的大小
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

Private Sub picToolBar_Click()
    Me.picToolBar.ZOrder
End Sub

Private Sub TabBar_TabClick(Frm As Form, Index As Integer)
    On Error Resume Next
    Frm.SetFocus                                                                    '点了TabBar之后让对应的窗口获得焦点
    Frm.SyntaxEdit.SetFocus
End Sub

Private Sub TabBar_WindowDropIn(Frm As Form, Index As Integer)
    On Error Resume Next
    Frm.SetFocus                                                                    '窗口拖进来后让对应的窗口获得焦点
    Frm.SyntaxEdit.SetFocus
End Sub

Private Sub TabBar_WindowDropOut(Frm As Form, Index As Integer)
    On Error Resume Next
    Frm.SetFocus                                                                    '窗口拖出去后让对应的窗口获得焦点
    Frm.SyntaxEdit.SetFocus
End Sub

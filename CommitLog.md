【日志】

# 2019.7.21

大幅删改ListView的代码。由于有皮肤控件帮忙，ListView可以更优雅的实现，于是去除所有不必要的控件和代码。另外把该控件的hWnd变量重命名为了lvHwnd。~~mmp是谁这样子起变量名的！！！让我发现不打死他！好像是自己起的哦...~~

把代码框左边的侧边栏去掉，并换成图片框控件。因为这个天杀的代码框居然没有提供获取断点的接口！！！我去年买了个表看样子要自己实现下断点的功能了。

优化用户体验，包括显示frmCreateOptions的时候让edProjectName获取焦点，自动生成WinMain代码的时候自动#include <windows.h>

为DarkVScrollBar控件的Public Property Let BarHeight添加了On Error Resume Next，因为有一定的出错几率。

TabBar现在会为子窗体添加WS_CHILD样式，使母窗体不失去焦点。

运行前的保存提示考虑工程文件。

gdb在附加进程前先发送`file 【待调试进程】`加载符号和`set pagination off`关闭gdb的"Type to continue, or q to quit"消息。

在主窗体的QueryUnload事件中添加隐藏菜单的代码，因为发现有一定的几率即使主窗体关闭后菜单也没被隐藏。

调整主窗体的QueryUnload事件中恢复窗口子类化的顺序，因为有时候会取消掉窗口关闭。

添加保存专用的代码文件信息结构和工程文件结构。

# 2019.7.20

把DarkComboBox的Image控件换成PNG控件，因为Image控件的颜色会被皮肤控件影响。

为TreeView添加了TVS_HASLINES样式。

把TabBar控件的SetFocus方法重命名为SwitchTo。

为TabBar控件添加SwitchToByForm方法，使TabBar能够切换到指定的窗口。

为TabBar控件添加UpdateCaptions方法，使TabBar能够更新所有窗口的标题。

添加了Win32Api.tlb文件到目录里。

clsPipe.cls里DosOutput的EndingStr参数说明之前忘记写了，现在补上。

把所有字符串常量都换成了变量，为之后切换语言的功能做准备。

把加载语言、显示Logo的代码放到了frmMain的Initialize事件中，因为语言必须比用户控件更早加载。

在创建工程后把工程文件标记为已更改。

把工程资源管理器里的“工程”替换成了当前工程的名称，并支持重命名。重命名文件的时候会自动选择“.”前面的字符串。

frmCreateOptions里的edProjectName文本框响应回车键。

frmCreateOptions在Load的时候不关闭皮肤，改成显示选择目录对话框的时候才关闭皮肤，选择完目录之后就立即恢复皮肤。

_保存工程的代码尚未编写！_

处理了gdb附加进程失败的情况。

处理了加载皮肤失败的清空。

为LoadLanguage函数添加了一个可选参数LoadMenuTextOnly，因为加载语言需要分成两步进行，首先在frmMain的Initialize事件中加载各种字符串，然后再在frmMain的Load事件中加载所有菜单语言。

# 2019.7.12

拖延了好几天（啊啊啊最近好忙），终于merge了404的pr...感谢404~

把按钮控件的AutoRedraw设置成True，避免按钮闪烁。

TabBar控件添加RemoveFormByForm方法。

# 2019.7.7

为DarkTitleBar添加了MinVisible和MaxVisible属性，可以选择隐藏最大、最小化按钮。如果有点Bug，有时候运行之后最大、最小化按钮就会自己隐藏，不知道为啥，也修不好，所以干脆运行时用代码设置算了。

完善DarkTreeView，有时使用SendMessageA没有以ByVal传参数，导致执行失败。现在已改成以ByVal方式传参数。

为DarkTreeView的UserControl_Resize添加修改树视图颜色的代码，使树视图不会被该死的皮肤控件改成难看的颜色。

修复TabBar的WindowDropOut事件不被触发的问题。

添加了全局变量IsExiting，该变量在退出时被设为True，一些窗体（如代码框）在关闭时检测到该变量为True时才会关闭，否则只是隐藏。

把frmCreate的大部分代码放到frmCreateOptions，因为frmCreateOptions会提供选项给用户设置，包括工程名称、路径等选项。

改进一些用户体验，如把MsgBox函数改成无皮肤的NoSkinMsgBox函数、frmCreate frmCreateOptions响应Esc键等、代码窗口拖出拖入的时候会获取焦点。

修复frmMain的Enabled设置成False时能调整大小的问题。（因为忘记设置frmMain.DarkWindowBorderSizer.Bind属性）

编写了frmCreateOptions窗体的代码。

添加frmMain的mnuSave_Click过程，该过程遍历CurrentProject.Files数组并保存所有标记为未保存的代码文件。

把编译相关的代码中有关部分当前工程的路径和名称。

把显示启动界面的过程单独写成一个过程ShowStartupPage。

为frmSolutionExplorer里的树视图添加双击代码，这部分代码会通过搜索TvItemBinding数组中匹配的树视图列表项句柄，从而确定列表项对应的文件序号，使列表项双击的时候会显示对应的代码（可能存在Bug）。

现在不直接操作frmCodeWindow，而是通过CreateNewCodeWindow函数来创建一个新的frmCodeWindow，再对创建的对象进行操作。因为CreateNewCodeWindow函数会往CodeWindows集合中添加新加的frmCodeWindow对象，使SourceFile结构中的TargetWindow能与其绑定。

编写ShowSave和ShowOpen函数，分别用来显示保存和加载通用对话框。通用对话框疑似会遭到皮肤重绘毒手，但是编译之后似乎又不会。所以暂时先不管吧。

初步定义工程各文件的信息结构雏形，包括工程主文件、工程代码文件等。

# 2019.7.3

移除了不同模块重复声明的一些API和函数，并把重复的操作改成子过程，让代码没这么臃肿。

添加了DarkTreeView控件，TreeView纯API创建的，一是不想又加多个ocx，二是VB6的TreeView控件会被皮肤控件画坏掉，用不成。由于TreeView是自制的，一定有许多功能，以及稳定性上不及VB6的TreeView控件。如果有发现我的TreeView有Bug，或者有什么需要改进的地方请告诉我。

我自制的TreeView控件的事件触发跟别的用户控件不一样，他的事件是直接通过子类化触发的，也就是说控件里面声明的Event实际上只是空壳。因为我技术实在欠佳，没有好的方式从子类化触发事件。我看到很多高手都是用Thunk的方式来写用户控件子类化的，不过我的水平远不及他们，实在是写不出o(╥﹏╥)o

在frmSolutionExplorer中添加了TreeView的所有事件。

优化了cmdNewPlainCpp_Click中代码的执行顺序，并在添加了操作TreeView的代码。

在OutputLog过程中添加了把光标移到结尾。

把ObjPtr(Me)改成了ObjPtr(WindowObj)，虽然效果一样，不过我认为这样更利于理解代码。

# 2019.7.2

为TabBar控件的Resize事件加了On Error Resume Next，因为有时候调整大小的时候会出错。

添加了管道类，用来获取DOS程序的输出。

添加检测指定进程是否存在的函数：ProcessExists

为frmOutput添加了文本框和OutputLog过程。

为frmPopupMenu的AddItems过程加了On Error Resume Next，暴力解决烦人的“表达式太复杂”错误（编译之后不会这样）。

修复主页面的排版，预留空位给工具栏。

主程序添加“运行”菜单的代码，尝试使用g++进行编译代码框里的代码，并使用管道获取其输出；并运行编译后的程序，用gdb进行附加。目前效果良好，不过尚未编写gdb调试相关的代码。想要试这个功能：新建空白C++程序→打代码→调试菜单→运行。

# ~~2019.6.31~~ 2019.7.1

优化窗口子类化，使其能适用于不同的窗口。窗口先用ObjPtr记录自身Object的地址，子类化通过获取这个Object来卸载窗体。

改进了ImageButton，现在支持图片加文本。

采用优化后的窗口子类化，修复了代码窗体最大化时全屏，以及代码窗体不能通过任务栏关闭的问题。

# 2019.6.30

主工程:

菜单位图现支持PNG。404可以尽情发挥了。

添加资源文件编译需要用到的头文件，路径为GCC\include\res。

极度（找不到适合的词汇）感谢404！！！~~抛弃~~搁置MuingIII，任劳任怨地帮我弄了这个漂亮的TabBar！！

IceControls:

添加了CtlBasic类和IceWindow类（没写多少）。

debug_build.bat:

修改了编译命令，现在会先编译资源文件，再编译exe。

# 2019.6.29

添加了GCC（千万别被上一个commit吓到！绝大部分都是GCC里的头文件！不是我写的2333333）

把ImageButton里的图片改成了PNG控件。

代码框显示了侧边栏，可以按那里下断点。（因为修复了皮肤样式）

添加了大部分需要用到的窗口，不过几乎什么都还没弄。

添加了创建工作区的代码。

主窗口里的客户区里面又多了一个窗口客户区，Pane在客户区里工作，而其他的窗口则在窗口客户区里工作。

添加了项目标识，用来记录当前工程是什么类型。

决定先弄“空白C++文件”的创建，应该不会太难。

菜单项一开始一些会禁用掉，之后会随不同工程的类型而启用不同的菜单项。

# 2019.6.27

创建了主项目工程。

添加了代码框窗口。

去掉了昨天所说的“遮挡用PictureBox”，因为通过修改样式文件修复了风格不一致的问题。

修复了子窗口调整大小不正确的问题。

更改DarkEdit的字体颜色，使其更容易阅读。

精简了一些不必要的样式文件。

应404要求，在工程目录中加入了所有用到的OCX文件。

修复了子窗口最大化后，主窗口大小改变时子窗口大小不随之改变的问题。

修复了窗口不能通过任务栏的右键菜单关闭的问题。

添加了菜单项。

添加了资源文件，该资源文件用于存档字串表，因为软件语言将会有中文和英文。以后程序里的字符串需要从资源文件里读取。

感谢404新添加的一些图标。

# 2019.6.26

修好了窗体最大化的时候全屏的问题。

创建了IceControls工程。该工程是整个项目的核心，负责为每个控件的功能提供接口，并处理每个控件的事件。头文件IceControls.h将采用header-only的方式编写。

添加了代码框。为了让代码框的滚动条能符合风格真的花了好多功夫...好辛苦啊 最后还用PictureBox遮挡了一下改不了的地方（23333）最后的效果还是可以的。

# 2019.6.25

加上了Docking Pane。问题：Pane的控制按钮风格不一致...尚不知道如何修复。

目前在XP上也能运行。

哦对，感谢404再次为UI做出的贡献！Pane控件的配色多亏他了。

# 2019.6.23

再次优化gdb附加进程的整个流程，现在流程如下：
1. 运行并挂起待调试进程；运行gdb
2. `file 【待调试进程】`
3. `set pagination off`
4. `attach 【待调试进程PID】`
5. `continue`

（代码懒得更新了，不过之后会按照这个流程进行）

下午弄了一个窗体的雏形，以后**可能**用这样的窗体。潜在问题：不知道自己以前弄的轮子是否可靠... 但愿可靠吧（23333）
欢迎就UI提出更改建议和意见！

感谢404为UI进行的改良！

# 2019.6.22

昨天下午到现在尝试使用管道来获取gdb的输出，经过许多尝试之后总算有点眉目了。代码请见“gdb管道调试测试”文件夹。
一些总结：
1. 看上去管道获取gdb输出是个不错的方式，之后拖控件大法的调试功能就打算这样弄了，通过分析gdb的输出来获取需要的信息
2. “中断程序”功能尚未完善，打算改良运行的方式：先用`CreateProcess`运行并挂起(`CREATE_SUSPENDED`)待调试的程序，然后让gdb附加到这个进程上。这样就可以得到待调试进程的PID和进程句柄，从而使用`DebugBreakProcess`使其中断
3. `NtSuspendProcess`虽然能让待调试进程挂起，但是gdb却仍然认为他在运行，故不采用此方式
4. 尚未试过gdb附加调试进程，之后会试试

潜在问题：
进入gdb之后需要使用`set pagination off`来关闭gdb的"Type <return> to continue, or q <return> to quit"消息

晚上对管道调试的机制进行了一些改进，现在已经能够基本达到我期望的标准，运行一个程序并对它进行调试。
一些总结：
1. 早上所提到的“中断程序”功能现在已经基本完善，并确实是像上面总结的第2点的思路做的
2. gdb附加调试进程很简单，只需要`attach 【进程ID】`即可

潜在问题：
调试完成之后可能会有残留的gdb进程或者待调试程序的进程，需要手动关闭。之后会想办法处理这些问题。

# 2019.6.21

今天早上又研究了一下gdb的调试命令，这些应该能用上：
1. `set variable 【变量名】 = 【新值】` 把指定变量的值更改成【新值】
2. `disassemble 【地址】` 在指定的地址反汇编
3. `continue` 继续执行
4. `tbreak` 设置一个临时断点，可以和`jump 【行号】`配合使用
5. `next` 逐过程执行； `step` 逐语句执行； `nexti` 逐个汇编指令执行
6. `finish` 执行到返回
7. `kill` 杀掉

# 2019.6.20

项目正式开坑！<br>
花了一个下午研究如何用GCC编译带有资源文件的exe。没想到原来没有想象中这么难。大致的步骤如下：<br>
1. 编写好资源文件脚本文件（Resource Script, *.rc）<br>
2. 使用windres对rc进行编译。windres会对rc文件里面涉及到的资源进行收集，所以尽量把里面涉及到的资源与rc文件放在同一目录下。编译命令：`windres --include-dir="【Include路径】" 【输入rc文件】 -O coff -o 【输出res文件】`<br>
3. 使用g++对源码和res文件进行编译。编译命令：`g++ 【源码路径】 【res路径】 -o 【输出exe路径】` 注意：【源码路径】不需要包括.h文件<br>

晚上研究了一下gdb的使用。如果我想要调试某个exe，那么使用g++编译它的时候要加上`-g`参数，使编译的exe能被gdb调试。<br>
接着试了一下gdb的命令行参数，大概是这样使用的（如有错请纠正，毕竟我是今天才开始学习使用的）：`gdb -exec="【需调试程序的路径】" -q -nw` 其中`-q`让gdb不要输出版本信息，`-nw`让gdb使用命令行界面。至于调试命令尚未研究，不过找到个文档，晚点看看：http://condor.depaul.edu/glancast/373class/docs/gdb.html

又研究了好一会儿，发现上面的命令行参数中`-exec="【需调试程序的路径】"`这一部分是多余的，因为即使在命令行中指定了需调试程序的路径之后也要指定。然后我熟悉了一下gdb的一些基本命令：
1. `file "【需调试程序的路径】"` 指定需要调试的程序，并加载符号（symbols）。注意路径里面的右斜杠“\”需要被替换成左斜杠“/”
2. `b 【函数名 或 行号】` 往指定的位置下断点
3. `d` 删除所有断点
4. `p 【表达式】` 非常好用！计算指定的表达式。如`p &a`能显示出变量a的地址
5. `info locals` 显示出本地变量
6. `info stack` 显示堆栈
7. `ptype 【表达式】` 同样非常好用！获取指定表达式的类型。如`ptype a`就能显示出变量a的类型（int等）。<br>

还有很多很多功能不一一罗列了。不得不感叹gdb真的十分强大！

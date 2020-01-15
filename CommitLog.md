【日志】

# 2020.1.15

修复标题栏图标双击不能关闭窗口的问题：原来是弹出菜单挡住了鼠标按键。把弹出菜单的位置改成了标题栏下方。

重写frmCallStack里面的GetCallStack()，现在（大概）能准确获取gdb的输出。~~看不懂原来的代码在干什么，大概脑袋抽筋。~~ 把原来的代码都注释掉了，不知道以后会不会再用上。

原来404修改创建工程窗口的时候留下了一个bug，如果通过主窗口的菜单创建工程的话主界面的frmCreate窗口不会卸载掉。修好了该bug。

增加了弹出菜单对键盘快捷键的支持（例如“打开(&O)”就能按O来触发）。

把一些字符串常数替换成了字符串变量。

对工程资源管理器的弹出菜单做了多语言支持。

（啰嗦话）原来我有近两个月没commit过了... 因为前段时间要准备面试，挺忙的。面试完之后就压力大减了 :) 写代码也开(tong)心(ku)点啦

# 2019.10.27

把检查文件名是否合法的过程单独写成一个函数`CheckInvalidFileName`，因为不同的位置（比如重命名、创建文件）会用到这个过程。

结构更改：新增文件夹信息结构`ProjectFolderStruct`；为代码文件信息结构`SourceFileStruct`添加`FolderIndex`属性，移除`IsHeaderFile`属性；为工程文件结构`ProjectFileStruct`添加`Folders()`数组，用来记录工程里面的所有文件夹；为树视图列表项绑定结构`TvItemToFileIndex`添加`IsFolder`属性，用来标记树视图列表项对应的项目是否为文件夹。

frmMain的mnuRun_Click过程中判断文件是否为头文件的思路由检查文件的`IsHeaderFile`属性改为检查文件的扩展名是否为h或者hpp。

frmMain的tmrCheckProcess_Timer过程中添加了检测并处理gdb进程意外退出的代码。

修改frmPopupMenu的PopupNewMenu过程，把新窗体的NoWhiteList属性设置为True，修复了Pane浮动窗口中弹出菜单的二级菜单在失焦后不会自动隐藏的Bug。

修改modTreeViewProc中树视图控件的子类化，处理其WM_CTLCOLOREDIT消息，让树视图编辑标签时的文本框的颜色更加友好。

修改modTreeViewProc中树视图的文本框的子类化，处理其WM_KEYDOWN消息，分别处理了方向键和Ctrl+A快捷键。

为工程资源管理器的树视图添加上下文菜单，并实现了添加文件夹、添加文件、重命名文件夹、重命名文件、用文件浏览器打开路径等过程。

为工程资源管理器的树视图添加菜单键（VK_APPS）的响应。（虽然可能根本没人会发现...）

# 2019.10.21

抱歉这么久没有commit，因为这段时间真的好忙。

改进DarkMenu的PopupMenu过程，把它的X、Y参数设置为可选的，默认弹出位置是鼠标指针的位置。在一些用户控件（如DarkTitleBar）对这个过程的调用也随之简化。

为frmPopupMenu添加NoWhitelist变量，并调整失焦检测计时器的代码，允许在类名为`XTPDockingPaneMiniWnd`的窗口上失去一次焦点，解决工程资源管理器在浮动的时候不能弹出菜单的问题。

为DarkTreeView控件添加HitTest函数，用来获取指定坐标位置对应的列表项。

把ImgOptionBox的Focused、Content属性的参数的龌龊至极的变量命名换成好看点的单词。

调整frmCreateOptions的“浏览”文本框和按钮的位置，让按钮不会遮住文本。

修改frmCreateOptions中创建工程的代码，现在创建工程的时候会自动创建工程文件和cpp文件。创建之前会检测是否有重名文件。这样可以解决创建工程之后不能重命名的问题。

修改frmCreateOptions中创建工程的代码，现在创建工程的时候不会在工程资源管理器的树视图里自动创建“源文件”节点。

完善一些窗体在加载时自动根据语言设置控件文本的代码。

为工程资源管理器窗口的树视图添加上下文菜单，不过代码尚未完善。

整理了一下图标的文件夹。

# 2019.9.1

感谢404为新建对话框的UI的改进。

弄好了搞乱掉的.gitignore。

删掉了一些不需要的文件。

现在frmCreate响应Esc键的时候会考虑是不是有标题栏。有的话才响应。

添加：当设置ImgOptionBox的Focused属性时会RaiseEvent Click

让代码的某些位置变得优雅一点点。

~~404要开学了。~~

# 2019.8.29

处理了一下新版gdb获取本地变量信息的问题。新版gdb的输出实在太令人头疼了，比如`<incomplete sequence \214>`这种东西。我实在不知道该怎么写代码处理了，请原谅我。

在Form_QueryUnload中添加禁用计时器的代码，防止gdb管道关闭之后仍然在获取管道内容。

重新添加了所有文件，解决了.gitignore无效的问题。以后就不会受到.vbw之类文件的影响了。

我觉得写这东西真的好累。一点都不快乐。原谅我过很长时间才push一点点东西。

# 2019.8.25

添加ByteArrayConv函数，替代VB6的StrConv的把字节数组转成字符串的功能。

移除了gcc。改成用户自行安装，并选择安装路径。（如需测试：修改frmMain的Form_Load中`GccPath`和`GdbPath`两个变量）

为断点列表添加按下显示详细信息的功能。

处理不同版本的gdb的路径字符问题（旧版是“/”，新版是“\”）。

为frmCreateOptions里的文本框响应Ctrl+A快捷键。

为本地变量列表响应鼠标双击事件。

添加ProcessExitedHandler过程，以处理调试进程退出之后的收尾工作。

为frmMain的ClearDebugWindows过程添加一个ClearBreakpoints参数，因为有时候清空调试窗口信息的时候不需要清空断点列表里的信息。

修改frmMain中运行部分的代码，让输出更加详细。

修改frmMain中运行部分的代码，在向第一次gdb发送“continue”命令之后使用`ResumeThread`来恢复主线程运行。对于旧版gdb，continue命令能让进程恢复运行；但是对于较新的版本，需要`ResumeThread`才可以恢复进程运行。

处理了新版gdb调试进程退出的输出。

# 2019.8.13

修改了DarkButton的颜色，使其跟DarkImageButton的颜色一致。

新添加函数StrConvEx，使用WideCharToMultiByte来替代VB6自己的StrConv，修复了在英文系统上中文会变成问号的问题。

修复了按下TabBar标签页的时候代码框不会获取到焦点的问题。

为调用堆栈的ListView添加了鼠标按键处理，按下鼠标按键能看到调用堆栈信息。

修复了frmMain的mnuRun中一处MsgBox没有改用NoSkinMsgBox的问题，导致弹出的消息框非常非常非常非常难看。

添加DestroyToolTip过程，用来在程序退出前关闭掉工具提示文本窗口以释放资源。

# 2019.8.12

经过n（n≥5）次拖延后终于添加了.gitigore文件，忽略掉.vbw文件和Vb_autoBak文件夹。

修复了DarkImgeButton控件颜色更改速度不一致的问题（鼠标移上去的时候颜色变亮得快，移出去的时候却变暗得慢）。

修复了DarkListView的严重错误，由于函数忘记返回数值导致控件事件不能和消息处理正确的绑定。

为DarkListView添加hWnd属性，返回用户控件的窗口句柄。

为frmBreakpoints添加ClearEverything过程，用于清空断点列表中的地址。这部分代码本来放在frmMain的ClearDebugWindows过程中，但是为了让风格一致，做了此修改。

编写了frmCallStack里的代码，能够对gdb的输出进行解析，并把调用堆栈的信息显示出来。写了frmLocals的代码之后觉得写这个窗口的代码轻松不少。

为一些有机会出错的过程添加了On Error Resume Next，在编译前应该去掉这些行的注释，尽量避免程序崩溃。（我已经尽量对所有可能出现的情况进行了考虑，但是恐有遗漏之处，考虑有不周到的地方，所以为了保险起见，还是添加这句）

为frmLocals添加ExpandItem过程，把展开列表项的代码单独写到一个过程里。让

改善了frmLocals的ArrayParser，处理了一开始查找字符串的时候找不到“"”的情况。

在frmLocals的Form_Load事件中初始化VarNodes数组，否则很大几率会导致编译的EXE未响应（即使已经加了On Error Resume Next）。

为frmLocals添加工具提示文本，在按下列表项之后显示其信息。

为frmLocals的ListView响应键盘的左、右方向键。

优化frmLocals的picSelMargin_MouseDown事件代码，把计算节点层数的代码移到了If分支里面，避免无谓的计算。

把窗体加载的时候一些NoSkinMsgBox换成了MsgBox，以及做了一些字符串常量和字符串变量的更换，因为考虑到窗体还没初始化完成，一些字符串资源尚未加载完成的情况。

把DockingPane创建的窗口名从常量改成了变量。

添加modTooltips.bas，工具提示文本模块。

# 2019.8.8

应404要求改了下DarkImageButton的鼠标移上去的颜色，使颜色更加明显一点。

去掉了DarkListView的WS_BORDER样式，因为他的边框在加载皮肤之后变成了黑框，不是很好看。

为一些更改窗口子类化的地方加了“ToDo”标志，方便之后修改。

clsPipe在关闭管道（CloseDosIO()）的时候会发送退出命令，让gdb退出，大大减小程序退出之后gdb仍在运行的几率。

为clsPipe的DosOutput函数添加一个可选的超时参数，如果传了这个参数，执行超时函数就会返回。

编写了frmLocals里的代码，能够对gdb的输出进行解析，并把本地变量的信息显示出来。这个窗口的代码真是写得我天昏地暗！！！呕心沥血！！！

把frmMain的GdbPipe改成了Public的，不要Private WithEvents了。因为别的窗体也需要用到它。

把清空调试窗口信息的代码单独写成了一个过程。

在编译前提示是否保存的时候添加了取消的选项，按下取消的时候会取消掉编译操作。

建立gdb管道之后发送`set print repeats 0`给gdb，关闭了gdb对于重复的数组元素的“<repeats n times>”输出。

等待附加进程的代码添加了超时，比较不优雅地解决了有时启动的时候会卡死的情况。

在frmMain的Form_QueryUnload事件中添加了关闭管道的代码，减小程序退出之后gdb仍在运行的几率。

优化tmrCheckProcess_Timer里的代码，防止获取gdb断点信息失败之后导致这里的代码出错。

处理了gdb输出`Program exited normally.`的消息。这种情况带代表进程返回了0。

修复了frmPopupMenu的菜单项有时候图标显示不正确的问题。

去掉了TabBar自动为窗体添加WS_CHILD样式，因为这样虽然可以让主窗体不失去焦点，但是有时候会有奇奇怪怪的问题，比如文本框经常失焦。

# 2019.7.29

为ListView控件添加Click事件（NM_CLICK）；修改DoubleClick事件的处理方式（WM_xBUTTONDBLCLK -> NM_DBLCLK）；添加GetItemChecked函数的代码，用以获取列表项是否已勾选。

为管道类添加ClearPipe函数，该函数使用ReadFile读取管道内的内容以清空管道。有时候分析gdb输出的时候会分析到之前的无用输出，因此添加该函数。

为代码窗口添加断点、禁用的断点和当前行的图片，供之后使用。

现在更改断点也视为更改了文件。

鼠标移动到断点栏上面会显示对应的端点的信息。

修复有时候菜单图标没绘制出来的问题。

把frmSolutionExplorer中SolutionTreeView_DoubleClick的代码弄得优雅一点。

添加运行、中断、停止的菜单图标。（感谢404帮忙绘制）

添加GdbBreakpointMapInfo用户类型，用来把gdb里面的断点序号跟不同文件里面的断点映射起来。

frmMain添加CurrState全局变量，用来记录当前的调试状态。

frmMain的mnuRun_Click：按下之后先检查CurrState，如果是中断状态的话就向gdb发送继续运行命令。

修复frmMain的mnuRun_Click中对重名EXE文件的检测，原来是ExePath还没赋值就用Dir去检测他了。

改善frmMain的mnuRun_Click代码排版，看起来似乎舒服多了？（x

在frmMain的mnuRun_Click中添加使用gdb下断点的代码。这部分的代码写得好辛苦，总结一下：
1. 使用DosInput的时候命令后面要添加换行符！否则命令就不执行了... 好几次都栽在这个坑里。
2. 断点列表里的最后一个元素是没有用到的。
3. gdb的输出需要逐行分析，否则直接进行分析会很乱、很复杂。而一行行拆开分析就好很多了。
4. 每次使用DosInput往gdb发送命令时应该先用ClearPipe清理管道，防止把上次的命令输出也一并分析了。

使用frmCheckProcess来定时获取gdb是否有输出内容，有的话对其分析。分别处理了断点命中消息和程序退出消息。

# 2019.7.25

为ListView控件添加GridLines和CurrExStyle属性，并改进了该控件调整样式的方式（使用CurrExStyle变量而不是直接用常数值更改样式，能够使控件的多种样式能同时使用）。

为ListView控件添加SetItemChecked方法和GetItemChecked方法（忘记编写了，晚点补上2333），用来获取列表项是否勾选。

为代码窗体添加断点相关的代码，如添加断点、绘制断点、文本更改时更改有关的断点、文本删除时删除断点等。

刷新断点的方式比较不优雅，是靠拦截代码框的WM_PAINT消息，然后做标记，告诉计时器断点需要刷新。暂时没有找到更优雅的方法。

调整代码窗体的Form_Load中的代码顺序，修复界面布局不能正确计算。

新建项目窗体：增加对工程名称、工程路径的命名检测。不允许特殊字符及空路径。按下确定键之后不直接创建cpp文件，而是等用户手动保存之后才创建。这样大概能更好的避免用户手贱把重要文件覆盖掉吧...(雾

frmMain的mnuSave_Click中处理没有文件需要保存的情况。

frmMain在显示frmStartupLogo的时候会Refresh它，否则它的内容显示不出来。

修复菜单项左边的图标跟菜单项不对齐的问题。

frmSaveBox的lstFiles_Click中处理没有选择保存文件的情况。如果用户一个文件都没有选择，就不给按下“是”。

修改modConfig.bas里的代码排版...强迫症（x

# 2019.7.22

把DarkButton的AutoRedraw设置成True，大概可以减少闪烁吧。

把frmCreateOptions和frmSaveBox的标题栏改名，使他们不能被拖进TabBar。

frmMain添加IsSaveRequired函数，用来检查是否有文件未保存。

把frmMain保存的代码放到了frmSaveBox里，能够直观的显示所有需要保存的文件，并让用户可以自行选择保存的文件。

优化frmMain中mnuRun_Click的保存文件的代码，更加易于使用。

frmMain的mnuRun_Click添加检查同名exe文件的过程，遇到同名的exe文件时会提示用户。

frmMain的Form_QueryUnload中添加保存文件的代码。

添加断点信息结构。同时也为代码文件信息结构中添加断点信息。

把代码文件信息结构的用户类型命名改得好听一点。（x

添加GetFileName函数，用于从指定路径分隔出文件名（即最后一个“\”后面的文本）

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

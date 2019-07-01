/*
描述:	主程序
作者:	冰棍
文件:	IceControls.cpp
*/

#include "IceControls.h"
#include "resource.h"

HINSTANCE ProgramInstance;						//程序的实例句柄

IceWindow MainWindow(IDD_MAINWINDOW);

/*
描述:	获取程序实例句柄的接口，供IceControls.h使用
返回值:	程序实例句柄
*/
HINSTANCE GetProgramInstance() {
	return ProgramInstance;
}

void Form_Load() {
	MessageBox(MainWindow.GetHwnd(), "Hello", "Ha", 0);
}

/*
描述:	程序入口点
参数:	hInstance: 程序的实例句柄
.		hPrevInstance: 程序上一个的实例句柄。总是为NULL
.		lpCmdLine: 程序的命令行
.		nCmdShow: 窗口显示命令
返回值:	程序运行返回值
*/
int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	//重要！不要删除！-------------------------
	ProgramInstance = hInstance;								//记录程序的实例句柄
	//----------------------------------------
	MainWindow.Form_Load = (void*)Form_Load;
	MainWindow.Create();
	return 0;
}

/*
描述:	定义各种控件类
作者:	冰棍
文件:	IceControls.h
*/

#pragma once

#include <Windows.h>

/*
描述:	获取当前程序的实例句柄的接口
返回值:	当前程序的实例句柄
*/
HINSTANCE GetProgramInstance();

/* 描述: 所有窗体的通用操作，包括调整大小、激活、显示等 */
class CtlBasic {
private:
	HWND	CtlHwnd;							//记录窗体的hWnd

protected:
	void SetHwnd(HWND NewHwnd) {				//设置窗体的hWnd，用于绑定窗体时使用
		CtlHwnd = NewHwnd;
	}

public:
	/*
	描述:	获取控件的hWnd
	返回值:	控件的hWnd
	*/
	HWND GetHwnd() {
		return CtlHwnd;
	}

	/*
	描述:	获取控件的水平位置
	返回值:	控件的水平位置
	*/
	LONG GetLeft() {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		return rc.left;
	}

	/*
	描述:	设置控件的水平位置
	*/
	void SetLeft(LONG NewLeft) {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		MoveWindow(CtlHwnd, NewLeft, rc.top, rc.right - rc.left, rc.bottom - rc.top, TRUE);
	}

	/*
	描述:	获取控件的垂直位置
	返回值:	控件的垂直位置
	*/
	LONG GetTop() {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		return rc.top;
	}

	/*
	描述:	设置控件的垂直位置
	*/
	void SetTop(LONG NewTop) {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		MoveWindow(CtlHwnd, rc.left, NewTop, rc.right - rc.left, rc.bottom - rc.top, TRUE);
	}

	/*
	描述:	获取控件的宽度
	返回值:	控件的宽度
	*/
	LONG GetWidth() {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		return rc.right - rc.left;
	}

	/*
	描述:	设置控件的宽度
	*/
	void SetWidth(LONG NewWidth) {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		MoveWindow(CtlHwnd, rc.left, rc.top, NewWidth, rc.bottom - rc.top, TRUE);
	}

	/*
	描述:	获取控件的高度
	返回值:	控件的高度
	*/
	LONG GetHeight() {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		return rc.bottom - rc.top;
	}

	/*
	描述:	设置控件的高度
	*/
	void SetHeight(LONG NewHeight) {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		MoveWindow(CtlHwnd, rc.left, rc.top, rc.right - rc.left, NewHeight, TRUE);
	}
};

/* 描述: 窗体类，提供窗体相关的操作 */
class IceWindow : public CtlBasic {
private:
	int		ResID;												//对话框资源ID

private:
	/*
	描述:	对话框消息处理过程
	参数:	hWnd: 窗口句柄
	.		uMsg: 消息值
	.		wParam, lParam: 消息的附加信息
	返回值:	消息处理结果
	*/
	static INT_PTR CALLBACK DialogWindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
		IceWindow	*IceWindowObj = (IceWindow*)GetProp(hWnd, "ClassObject");			//读取窗口对应的IceWindow类

		switch (uMsg) {
		case WM_INITDIALOG:																//对话框初始化
			SetProp(hWnd, "ClassObject", (HANDLE)lParam);									//记录窗口对应的IceWindow类
			if (((IceWindow*)lParam)->Form_Load)											//如果有绑定Form_Load事件
				((void(*)())(((IceWindow*)lParam)->Form_Load))();								//调用绑定的Form_Load
			break;

			
		}
		
		return FALSE;																	//使用系统默认的处理方式
	}

public:
	//事件绑定
	void	*Form_Load;
	void	*Form_Activate;

	/*
	描述:	声明一个IceWindow，并初始化事件绑定
	参数:	DialogResID: 对话框资源ID
	*/
	IceWindow(int DialogResID) {
		ResID = DialogResID;
		Form_Load = NULL;
		Form_Activate = NULL;
	}

	/*
	描述:	创建指定资源ID的对话框
	*/
	void Create() {
		DialogBoxParam(GetProgramInstance(), MAKEINTRESOURCE(ResID), NULL, DialogWindowProc, (LPARAM)this);
	}
};

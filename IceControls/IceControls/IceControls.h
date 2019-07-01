/*
����:	������ֿؼ���
����:	����
�ļ�:	IceControls.h
*/

#pragma once

#include <Windows.h>

/*
����:	��ȡ��ǰ�����ʵ������Ľӿ�
����ֵ:	��ǰ�����ʵ�����
*/
HINSTANCE GetProgramInstance();

/* ����: ���д����ͨ�ò���������������С�������ʾ�� */
class CtlBasic {
private:
	HWND	CtlHwnd;							//��¼�����hWnd

protected:
	void SetHwnd(HWND NewHwnd) {				//���ô����hWnd�����ڰ󶨴���ʱʹ��
		CtlHwnd = NewHwnd;
	}

public:
	/*
	����:	��ȡ�ؼ���hWnd
	����ֵ:	�ؼ���hWnd
	*/
	HWND GetHwnd() {
		return CtlHwnd;
	}

	/*
	����:	��ȡ�ؼ���ˮƽλ��
	����ֵ:	�ؼ���ˮƽλ��
	*/
	LONG GetLeft() {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		return rc.left;
	}

	/*
	����:	���ÿؼ���ˮƽλ��
	*/
	void SetLeft(LONG NewLeft) {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		MoveWindow(CtlHwnd, NewLeft, rc.top, rc.right - rc.left, rc.bottom - rc.top, TRUE);
	}

	/*
	����:	��ȡ�ؼ��Ĵ�ֱλ��
	����ֵ:	�ؼ��Ĵ�ֱλ��
	*/
	LONG GetTop() {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		return rc.top;
	}

	/*
	����:	���ÿؼ��Ĵ�ֱλ��
	*/
	void SetTop(LONG NewTop) {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		MoveWindow(CtlHwnd, rc.left, NewTop, rc.right - rc.left, rc.bottom - rc.top, TRUE);
	}

	/*
	����:	��ȡ�ؼ��Ŀ��
	����ֵ:	�ؼ��Ŀ��
	*/
	LONG GetWidth() {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		return rc.right - rc.left;
	}

	/*
	����:	���ÿؼ��Ŀ��
	*/
	void SetWidth(LONG NewWidth) {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		MoveWindow(CtlHwnd, rc.left, rc.top, NewWidth, rc.bottom - rc.top, TRUE);
	}

	/*
	����:	��ȡ�ؼ��ĸ߶�
	����ֵ:	�ؼ��ĸ߶�
	*/
	LONG GetHeight() {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		return rc.bottom - rc.top;
	}

	/*
	����:	���ÿؼ��ĸ߶�
	*/
	void SetHeight(LONG NewHeight) {
		RECT rc;
		GetWindowRect(CtlHwnd, &rc);
		MoveWindow(CtlHwnd, rc.left, rc.top, rc.right - rc.left, NewHeight, TRUE);
	}
};

/* ����: �����࣬�ṩ������صĲ��� */
class IceWindow : public CtlBasic {
private:
	int		ResID;												//�Ի�����ԴID

private:
	/*
	����:	�Ի�����Ϣ�������
	����:	hWnd: ���ھ��
	.		uMsg: ��Ϣֵ
	.		wParam, lParam: ��Ϣ�ĸ�����Ϣ
	����ֵ:	��Ϣ������
	*/
	static INT_PTR CALLBACK DialogWindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
		IceWindow	*IceWindowObj = (IceWindow*)GetProp(hWnd, "ClassObject");			//��ȡ���ڶ�Ӧ��IceWindow��

		switch (uMsg) {
		case WM_INITDIALOG:																//�Ի����ʼ��
			SetProp(hWnd, "ClassObject", (HANDLE)lParam);									//��¼���ڶ�Ӧ��IceWindow��
			if (((IceWindow*)lParam)->Form_Load)											//����а�Form_Load�¼�
				((void(*)())(((IceWindow*)lParam)->Form_Load))();								//���ð󶨵�Form_Load
			break;

			
		}
		
		return FALSE;																	//ʹ��ϵͳĬ�ϵĴ���ʽ
	}

public:
	//�¼���
	void	*Form_Load;
	void	*Form_Activate;

	/*
	����:	����һ��IceWindow������ʼ���¼���
	����:	DialogResID: �Ի�����ԴID
	*/
	IceWindow(int DialogResID) {
		ResID = DialogResID;
		Form_Load = NULL;
		Form_Activate = NULL;
	}

	/*
	����:	����ָ����ԴID�ĶԻ���
	*/
	void Create() {
		DialogBoxParam(GetProgramInstance(), MAKEINTRESOURCE(ResID), NULL, DialogWindowProc, (LPARAM)this);
	}
};

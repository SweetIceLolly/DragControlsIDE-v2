/*
����:	������
����:	����
�ļ�:	IceControls.cpp
*/

#include "IceControls.h"
#include "resource.h"

HINSTANCE ProgramInstance;						//�����ʵ�����

IceWindow MainWindow(IDD_MAINWINDOW);

/*
����:	��ȡ����ʵ������Ľӿڣ���IceControls.hʹ��
����ֵ:	����ʵ�����
*/
HINSTANCE GetProgramInstance() {
	return ProgramInstance;
}

void Form_Load() {
	MessageBox(MainWindow.GetHwnd(), "Hello", "Ha", 0);
}

/*
����:	������ڵ�
����:	hInstance: �����ʵ�����
.		hPrevInstance: ������һ����ʵ�����������ΪNULL
.		lpCmdLine: �����������
.		nCmdShow: ������ʾ����
����ֵ:	�������з���ֵ
*/
int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	//��Ҫ����Ҫɾ����-------------------------
	ProgramInstance = hInstance;								//��¼�����ʵ�����
	//----------------------------------------
	MainWindow.Form_Load = (void*)Form_Load;
	MainWindow.Create();
	return 0;
}

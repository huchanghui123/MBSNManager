#pragma once
#include <afxwin.h>
class MyCEditEx :
	public CEdit
{
	//����������ַ�
	BOOL PreTranslateMessage(MSG* pMsg);
};


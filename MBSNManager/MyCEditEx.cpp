#include "pch.h"
#include "MyCEditEx.h"

//ֻ��������10�����ַ���
BOOL MyCEditEx::PreTranslateMessage(MSG * pMsg)
{
	if (WM_CHAR == pMsg->message)
	{
		if (pMsg->wParam == VK_BACK)//��Ӧback��
			return CEdit::PreTranslateMessage(pMsg);
		TCHAR ch = (TCHAR)pMsg->wParam;
		if ((ch >= '0' && ch <= '9'))
			return CEdit::PreTranslateMessage(pMsg);
		
		return TRUE;
	}
	return CEdit::PreTranslateMessage(pMsg);
}

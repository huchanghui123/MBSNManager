#include "pch.h"
#include "MyCEditEx.h"
#include "manager.h"

//ֻ��������10�����ַ���
BOOL MyCEditEx::PreTranslateMessage(MSG * pMsg)
{
	CString temp;
	if (WM_CHAR == pMsg->message)
	{
		 GetWindowText(temp);
		 //printf("%s", temp);
		//��Ӧback��
		if (pMsg->wParam == VK_BACK)
		{
			return CEdit::PreTranslateMessage(pMsg);
		}
		
		TCHAR ch = (TCHAR)pMsg->wParam;
		if ((ch >= 'a' && ch <= 'z'))
			return CEdit::PreTranslateMessage(pMsg);
		if ((ch >= 'A' && ch <= 'Z'))
			return CEdit::PreTranslateMessage(pMsg);
		if ((ch >= '0' && ch <= '9'))
			return CEdit::PreTranslateMessage(pMsg);
		
		return TRUE;
	}
	return CEdit::PreTranslateMessage(pMsg);
}

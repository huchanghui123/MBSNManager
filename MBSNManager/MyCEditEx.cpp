#include "pch.h"
#include "MyCEditEx.h"
#include "manager.h"

//ֻ��������10�����ַ���
BOOL MyCEditEx::PreTranslateMessage(MSG * pMsg)
{
	if (WM_CHAR == pMsg->message)
	{
		//��Ӧback��
		if (pMsg->wParam == VK_BACK)
		{
			return CEdit::PreTranslateMessage(pMsg);
		}
		if(addSnl<3)//ǰ��λ�����
		{ 
			return CEdit::PreTranslateMessage(pMsg);
		}
		else 
		{
			TCHAR ch = (TCHAR)pMsg->wParam;
			if ((ch >= '0' && ch <= '9'))
			{
				return CEdit::PreTranslateMessage(pMsg);
			}
		}
		
		return TRUE;
	}
	return CEdit::PreTranslateMessage(pMsg);
}

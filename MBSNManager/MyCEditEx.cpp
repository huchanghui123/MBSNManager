#include "pch.h"
#include "MyCEditEx.h"
#include "manager.h"

//只允许输入10进制字符串
BOOL MyCEditEx::PreTranslateMessage(MSG * pMsg)
{
	if (WM_CHAR == pMsg->message)
	{
		//相应back键
		if (pMsg->wParam == VK_BACK)
		{
			return CEdit::PreTranslateMessage(pMsg);
		}
		if(addSnl<3)//前三位随便输
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

#include "pch.h"
#include "MyCEditEx.h"
#include "manager.h"

//只允许输入10进制字符串
BOOL MyCEditEx::PreTranslateMessage(MSG * pMsg)
{
	if (pMsg->message == WM_KEYDOWN)
	{
		BOOL Control = ::GetKeyState(VK_CONTROL) & 0x8000;	
		if (Control && (pMsg->wParam == _T('v')|| pMsg->wParam == _T('V')))
		{
			this->Paste();
			return	TRUE;
		}
	}

	if (WM_CHAR == pMsg->message)
	{
		//响应back键
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
		if ((ch == '-'))
			return CEdit::PreTranslateMessage(pMsg);
		
		return TRUE;
	}
	return CEdit::PreTranslateMessage(pMsg);
}

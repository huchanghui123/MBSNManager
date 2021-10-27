#pragma once
#include <afxwin.h>
class MyCEditEx :
	public CEdit
{
	//¼àÌıÊäÈëµÄ×Ö·û
	BOOL PreTranslateMessage(MSG* pMsg);
};


#pragma once

typedef struct tagSNDATA {
	int sid;
	CString noStr, dateStr, modelStr, snStr, clientStr, saleStr;
}SNDATA, *LPSNDATA;

extern CString accessName;
extern CString accessPath;

#define WM_ONINIT_ACCESS WM_USER + 100
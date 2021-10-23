#pragma once


using std::vector;

typedef struct tagSNDATA {
	_variant_t sid, order, ordate, model, sn, client, sale;
}SNDATA, *LPSNDATA;

extern CString accessName;
extern CString accessPath;

#define WM_ONINIT_ACCESS WM_USER + 100
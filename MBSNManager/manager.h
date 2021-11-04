#pragma once

using std::vector;

typedef struct tagSNDATA {
	int sid;
	CString order, ordate, model, sn, client, sale;
}SNDATA, *LPSNDATA;

extern CString accessFile;
extern CString accessPath;
extern CString accessName;
extern CString tableName;
extern CString errorMsg;
extern int addSnl;


#define WM_CREATEDB_ACCESS WM_USER + 100
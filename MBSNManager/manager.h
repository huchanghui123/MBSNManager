#pragma once

#define WM_CREATEDB_ACCESS WM_USER + 100

using std::vector;

typedef struct tagSNDATA {
	int sid;
	CString order, ordate, model, sn, client, sale, status, typeSN;
}SNDATA, *LPSNDATA;

extern CStringArray mbTypsArr;
extern CString accessFile;
extern CString accessPath;
extern CString accessName;
extern CString tableName;
extern CString errorMsg;


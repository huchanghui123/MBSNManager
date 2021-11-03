#pragma once

#include "manager.h"

class ADOTools
{
public:
	ADOTools(void);
	~ADOTools(void);

	//连接对象智能指针，通常用于创建一个数据库连接或执行一条不返回任何结果的SQL语句
	_ConnectionPtr m_pConnection;
	//记录集对象智能指针，它对记录集提供了控制功能
	_RecordsetPtr m_pRecordset;

	BOOL CreateADOData(CString);
	BOOL OnConnADODB();
	vector<SNDATA> GetADODBForSql(LPCTSTR);
	vector<CString> GetADODBSNForSql(LPCTSTR);
	BOOL OnExecuteADODB(LPCTSTR);

	void GetDBTableNames();

	void ExitADOConn(void);
};


#pragma once
class ADOTools
{
public:
	ADOTools(void);
	~ADOTools(void);

	BOOL CreateAccessData(CString);
	//连接对象智能指针，通常用于创建一个数据库连接或执行一条不返回任何结果的SQL语句
	_ConnectionPtr m_pConnection;
	//记录集对象智能指针，它对记录集提供了控制功能
	_RecordsetPtr m_pRecordset;

};


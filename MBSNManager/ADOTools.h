#pragma once

#include "manager.h"

class ADOTools
{
public:
	ADOTools(void);
	~ADOTools(void);

	//���Ӷ�������ָ�룬ͨ�����ڴ���һ�����ݿ����ӻ�ִ��һ���������κν����SQL���
	_ConnectionPtr m_pConnection;
	//��¼����������ָ�룬���Լ�¼���ṩ�˿��ƹ���
	_RecordsetPtr m_pRecordset;

	vector<SNDATA> dbVector;

	BOOL CreateADOData(CString);
	BOOL OnConnADODB();
	vector<SNDATA> GetADODBForSql(LPCTSTR);


	void ExitADOConn(void);
};

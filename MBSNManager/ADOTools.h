#pragma once
class ADOTools
{
public:
	ADOTools(void);
	~ADOTools(void);

	BOOL CreateAccessData(CString);
	//���Ӷ�������ָ�룬ͨ�����ڴ���һ�����ݿ����ӻ�ִ��һ���������κν����SQL���
	_ConnectionPtr m_pConnection;
	//��¼����������ָ�룬���Լ�¼���ṩ�˿��ƹ���
	_RecordsetPtr m_pRecordset;

};


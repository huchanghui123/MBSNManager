#include "pch.h"
#include "ADOTools.h"
#include "manager.h"


ADOTools::ADOTools(void)
{
}

ADOTools::~ADOTools(void)
{
}


BOOL ADOTools::CreateADOData(CString fileName)
{
	if (PathFileExists(fileName))
	{
		AfxMessageBox(_T("���ݿ��Ѵ���!"));
		return FALSE;
	}
	else
	{
		ADOX::_CatalogPtr m_pCatalog = NULL;
		_bstr_t ConnectString = "Provider=Microsoft.ACE.OLEDB.16.0; Data Source = " +
			(_bstr_t)fileName;
		ADOX::_TablePtr m_pTable = NULL;
		try
		{
			//�������ݿ�
			m_pCatalog.CreateInstance(__uuidof(ADOX::Catalog));
			m_pCatalog->Create((_bstr_t)ConnectString);
			//�������ݿ�
			m_pConnection = _ConnectionPtr(__uuidof(Connection));
			m_pConnection->ConnectionTimeout = 20;
			m_pConnection->Open(ConnectString, "", "", adModeUnknown);
			
			//�������ݱ�
			//ID������OrderNo�����ţ�OrderDateʱ�䣬Model�ͺţ�SerialNo���кţ�Client�ͻ���Saleҵ��Status״̬��TYPE_SN�ͺ�_���к�(��Ϊ�յ�Ψһ�ֶΣ���Ϊ�в�ͬ�ͺ��������к��ظ�)
			m_pConnection->BeginTrans();
			_variant_t RecordsAffected;
			_bstr_t bstr1 = "CREATE TABLE SNTable";
			_bstr_t bstr2 = "(ID AUTOINCREMENT PRIMARY KEY, OrderNo Text, OrderDate Text, Model Text, SerialNo Text, Client Text, Sale Text, Status Text, TYPE_SN Text NOT NULL Unique)";
			_bstr_t CommandText = bstr1 + bstr2;
			//m_pConnection->Execute(CommandText, &RecordsAffected, adCmdText);

			struct _Recordset * _result = 0;
			HRESULT _hr = m_pConnection->raw_Execute(CommandText, &RecordsAffected, adCmdText, &_result);
			if (FAILED(_hr))
			{
				_com_issue_errorex(_hr, m_pConnection, __uuidof(_ConnectionPtr));
				m_pConnection->RollbackTrans();
			}
			else
			{
				m_pConnection->CommitTrans();
			}
			
			if ( m_pConnection->State == adStateOpen)
			{
				m_pConnection->Close();
			}

		}
		catch (_com_error &e)
		{
			CString errormessage;
			errormessage.Format(_T("%s"), e.ErrorMessage());
			errormessage.Append(_T("\r\n"));
			errormessage.Append(e.Description());
			AfxMessageBox(errormessage);
			return FALSE;
		}
	}
	return TRUE;
}

BOOL ADOTools::OnConnADODB()
{
	_bstr_t mysql = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" + (_bstr_t)accessFile;
	//_bstr_t mysql = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + (_bstr_t)accessFile;
	try
	{
		CoInitialize(NULL);
		m_pConnection = _ConnectionPtr(__uuidof(Connection));//�������Ӷ���
		//m_pConnection->ConnectionString = mysql;
		//AfxMessageBox(mysql);
		m_pConnection->ConnectionTimeout = 20;
		//adConnectUnsepecified(Ĭ��ֵ��ͬ������adAsyncConnect(�첽��
		//m_pConnection->Open(mysql, "", "", adConnectUnspecified);

		HRESULT hr = m_pConnection->raw_Open(mysql, (BSTR)"", (BSTR)"", adConnectUnspecified);
		if (FAILED(hr))
		{
			_com_issue_errorex(hr, m_pConnection, __uuidof(_ConnectionPtr));
		}

		m_pRecordset = _RecordsetPtr(__uuidof(Recordset));//������¼������
		GetDBTableNames();//��ʼ������
		if (tableName.GetLength()==0)
		{
			return FALSE;
		}
	}
	catch (_com_error &e)
	{
		CString errormessage;
		errormessage.Format(_T("%s"), e.ErrorMessage());
		errormessage.Append(_T("\r\n"));
		errormessage.Append(e.Description());
		AfxMessageBox(errormessage);
		return FALSE;
	}
	return TRUE;
}

vector<SNDATA> ADOTools::GetADODBForSql(LPCTSTR lpSql)
{
	vector<SNDATA> dbVector = {};
	
	if (m_pRecordset!=NULL)
	{
		//m_pRecordset->Open((_variant_t)lpSql, m_pConnection.GetInterfacePtr(),
		//	adOpenDynamic, adLockOptimistic, adCmdText);
		try
		{
			HRESULT hr = m_pRecordset->raw_Open((_variant_t)lpSql, (_variant_t)m_pConnection.GetInterfacePtr(),
				adOpenDynamic, adLockOptimistic, adCmdText);
			if (FAILED(hr))
				_com_issue_errorex(hr, m_pRecordset, __uuidof(_Recordset));
		}
		catch (_com_error &e)
		{
			CString errormessage;
			errormessage.Format(_T("%s"), e.ErrorMessage());
			errormessage.Append(_T("\r\n"));
			errormessage.Append(e.Description());
			AfxMessageBox(errormessage);
			return dbVector;
		}
	}
	
	while (!m_pRecordset->adoEOF)
	{	
		SNDATA data = {};
		data.sid = m_pRecordset->GetCollect(_T("ID")).intVal;
		data.order = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("OrderNo"));
		data.ordate = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("OrderDate"));
		data.model = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("Model"));
		data.sn = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("SerialNo"));
		data.client = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("Client"));
		data.sale = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("Sale"));
		data.status = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("Status"));
		dbVector.push_back(data);

		m_pRecordset->MoveNext();
	}
	
	if (m_pRecordset->State == adStateOpen)
	{
		m_pRecordset->Close();
	}

	return dbVector;
}

vector<CString> ADOTools::GetADODBSNForSql(LPCTSTR lpSql)
{
	vector<CString> snVector = {};
	m_pRecordset->Open((_variant_t)lpSql, m_pConnection.GetInterfacePtr(),
		adOpenDynamic, adLockOptimistic, adCmdText);
	if (m_pRecordset == NULL)
	{
		return snVector;
	}
	while (!m_pRecordset->adoEOF)
	{
		CString snStr = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("SerialNo"));
		snVector.push_back(snStr);

		m_pRecordset->MoveNext();
	}
	if (m_pRecordset->State == adStateOpen)
	{
		m_pRecordset->Close();
	}
	return snVector;
}

BOOL ADOTools::OnExecuteADODB(LPCTSTR lpSql)
{
	_variant_t RecordsAffected;
	try
	{
		m_pConnection->BeginTrans();
		//m_pConnection->Execute((_bstr_t)lpSql, &RecordsAffected, adCmdText);
		struct _Recordset * _result = 0;
		HRESULT _hr = m_pConnection->raw_Execute((_bstr_t)lpSql, &RecordsAffected, adCmdText, &_result);
		if (FAILED(_hr))
		{
			_com_issue_errorex(_hr, m_pConnection, __uuidof(_ConnectionPtr));
			m_pConnection->RollbackTrans();
		}
		else
		{
			m_pConnection->CommitTrans();
		}
	}
	catch (_com_error &e)
	{
		CString errormessage;
		errormessage.Format(_T("%s"), e.ErrorMessage());
		errormessage.Append(_T("\r\n"));
		errormessage.Append(e.Description());
		errorMsg = errormessage;
		//AfxMessageBox(errormessage);
		return FALSE;
	}

	return TRUE;
}

void ADOTools::GetDBTableNames()
{
	try
	{
			m_pRecordset = m_pConnection->OpenSchema(adSchemaTables);
			while (!(m_pRecordset->adoEOF))
			{
				//��ȡ���   
				_bstr_t table_name = m_pRecordset->Fields->GetItem("TABLE_NAME")->Value;

				//��ȡ�������        
				_bstr_t table_type = m_pRecordset->Fields->GetItem("TABLE_TYPE")->Value;

				//����һ�£�ֻ���������ƣ�������ʡ��
				if (strcmp(((LPCSTR)table_type), "TABLE") == 0) {
					tableName = (LPCSTR)table_name;
					//AfxMessageBox(tableName);
				}
				m_pRecordset->MoveNext();
			}
			if (m_pRecordset->State)
			{
				m_pRecordset->Close();
			}
		
	}
	catch (_com_error e)///��׽�쳣
	{
		CString errormessage;
		errormessage.Format(_T("�������ݿ�ʧ��!rn������Ϣ:%s"), e.ErrorMessage());
		AfxMessageBox(errormessage);
	}
}

void ADOTools::ExitADOConn()
{
	try
	{
		if (m_pConnection!=NULL && m_pConnection->State)
		{
			m_pConnection->Close();
		}
		m_pConnection = NULL;
		if (m_pRecordset != NULL && m_pRecordset->State)
		{
			m_pRecordset->Close();
		}
		m_pRecordset = NULL;
		
	}
	catch (_com_error e)
	{
		AfxMessageBox(e.ErrorMessage());
	}
}



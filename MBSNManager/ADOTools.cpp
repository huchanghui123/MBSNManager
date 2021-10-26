#include "pch.h"
#include "ADOTools.h"
#include "manager.h"


ADOTools::ADOTools(void)
{
}

ADOTools::~ADOTools(void)
{
}



BOOL ADOTools::CreateADOData(CString name)
{
	TCHAR filePath[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, filePath);

	CString fileName;
	fileName.Format(_T("%s"),filePath);
	fileName = fileName + _T("\\") + name+ _T(".accdb");

	if (PathFileExists(fileName))
	{
		AfxMessageBox(_T("数据库已存在!"));
		return FALSE;
	}
	else
	{
		ADOX::_CatalogPtr m_pCatalog = NULL;
		_bstr_t ConnectString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " +
			(_bstr_t)name + ".accdb;";
		ADOX::_TablePtr m_pTable = NULL;
		try
		{
			//创建数据库
			m_pCatalog.CreateInstance(__uuidof(ADOX::Catalog));
			m_pCatalog->Create((_bstr_t)ConnectString);
			//连接数据库
			m_pConnection = _ConnectionPtr(__uuidof(Connection));
			m_pConnection->ConnectionTimeout = 20;
			m_pConnection->Open(ConnectString, "", "", adModeUnknown);
			//创建表
			//m_pCatalog->PutActiveConnection(ConnectString);
			/*m_pTable.CreateInstance(__uuidof(ADOX::Table));
			m_pTable->PutName("SNTable");
			m_pTable->Columns->Append("OrderNo", ADOX::adBSTR, 0);
			m_pTable->Columns->Append("Date", ADOX::adBSTR, 0);
			m_pTable->Columns->Append("Model", ADOX::adBSTR, 0);
			m_pTable->Columns->Append("SnNo", ADOX::adBSTR, 0);
			m_pTable->Columns->Append("Client", ADOX::adBSTR, 0);
			m_pTable->Columns->Append("Sale", ADOX::adBSTR, 0);
			m_pCatalog->Tables->Append(
				_variant_t((IDispatch *)m_pTable));*/

			m_pConnection->BeginTrans();
			_variant_t RecordsAffected;
			_bstr_t bstr1 = "CREATE TABLE SNTable";
			_bstr_t bstr2 = "(OrderNo Text, Date Text, Model Text, SnNo Text, Client Text, Sale Text)";
			_bstr_t CommandText = bstr1 + bstr2;
			m_pConnection->Execute(CommandText, &RecordsAffected, adCmdText);
			m_pConnection->CommitTrans();

			accessPath = fileName;
			accessName = name + _T(".accdb");
		}
		catch (_com_error* e)
		{
			AfxMessageBox(e->ErrorMessage());
			return FALSE;
		}
	}
	return TRUE;
}

BOOL ADOTools::OnConnADODB()
{
	try
	{
		CoInitialize(NULL);
		m_pConnection = _ConnectionPtr(__uuidof(Connection));//创建连接对象
		m_pConnection->ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=test.accdb;";
		m_pConnection->ConnectionTimeout = 20;
		//adConnectUnsepecified(默认值，同步）和adAsyncConnect(异步）
		m_pConnection->Open("", "", "", adConnectUnspecified);
		m_pRecordset = _RecordsetPtr(__uuidof(Recordset));//创建记录集对象
	}
	catch (_com_error* e)
	{
		AfxMessageBox(e->ErrorMessage());
		return FALSE;
	}
	return TRUE;
}

vector<SNDATA> ADOTools::GetADODBForSql(LPCTSTR lpSql)
{
	vector<SNDATA> dbVector = {};
	m_pRecordset->Open((_variant_t)lpSql, m_pConnection.GetInterfacePtr(),
		adOpenDynamic, adLockOptimistic, adCmdText);
	if (m_pRecordset == NULL)
	{
		AfxMessageBox(_T("读取数据记录发生错误"));
		return dbVector;
	}
	
	while (!m_pRecordset->adoEOF)
	{	
		SNDATA data = {};
		data.sid = m_pRecordset->GetCollect(_T("ID")).intVal;
		data.order = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("OrderNo"));
		data.ordate = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("OrderDate"));
		data.model = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("Model"));
		data.sn = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("SerialNo"));
		data.client = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("Clinet"));
		data.sale = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect(_T("Sale"));
		dbVector.push_back(data);

		m_pRecordset->MoveNext();
	}
	
	if (m_pRecordset->State == adStateOpen)
	{
		m_pRecordset->Close();
	}

	return dbVector;
}

BOOL ADOTools::OnAddADODB(SNDATA snd)
{
	_variant_t RecordsAffected;
	TCHAR szSql[1024] = { 0 };
	_stprintf(szSql, _T("INSERT INTO SNTable(OrderNo,OrderDate,Model,SerialNo,Clinet,Sale) VALUES('%s','%s','%s','%s','%s','%s')"),
		(LPCTSTR)snd.order, (LPCTSTR)snd.ordate, (LPCTSTR)snd.model,
		(LPCTSTR)snd.sn, (LPCTSTR)snd.client, (LPCTSTR)snd.sale);
	try {
		m_pConnection->BeginTrans();
		m_pConnection->Execute((_bstr_t)szSql, &RecordsAffected, adCmdText);
		m_pConnection->CommitTrans();
	}
	catch (_com_error &e) {
		AfxMessageBox((LPCTSTR)e.Description());
		return FALSE;
	}

	return TRUE;
}

BOOL ADOTools::OnDelADODB(LPCTSTR lpSql)
{
	_variant_t RecordsAffected;
	try
	{
		m_pConnection->Execute((_bstr_t)lpSql, &RecordsAffected, adCmdText);
	}
	catch (_com_error &e)
	{
		AfxMessageBox((LPCTSTR)e.Description());
		return FALSE;
	}

	return TRUE;
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
	}
	catch (_com_error e)
	{
		AfxMessageBox(e.ErrorMessage());
	}
}



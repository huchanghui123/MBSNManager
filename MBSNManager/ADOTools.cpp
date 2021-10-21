#include "pch.h"
#include "ADOTools.h"
#include "manager.h"

ADOTools::ADOTools(void)
{
}

ADOTools::~ADOTools(void)
{
}



BOOL ADOTools::CreateAccessData(CString name)
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

		try
		{
			m_pCatalog.CreateInstance(__uuidof(ADOX::Catalog));
			m_pCatalog->Create((_bstr_t)ConnectString);

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



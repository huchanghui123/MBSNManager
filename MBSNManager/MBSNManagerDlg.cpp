
// MBSNManagerDlg.cpp: 实现文件
//

#include "pch.h"
#include "framework.h"
#include "MBSNManager.h"
#include "MBSNManagerDlg.h"
#include "afxdialogex.h"
#include "CreateData.h"
#include "ADOTools.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnAbout();
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CMBSNManagerDlg 对话框



CMBSNManagerDlg::CMBSNManagerDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_MBSNMANAGER_DIALOG, pParent)
	, addNumEdit(1)
	, mfNum(1)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMBSNManagerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, mList);
	DDX_Text(pDX, IDC_ADD_NUM, addNumEdit);

	ListView_SetExtendedListViewStyle(mList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
	DDX_Control(pDX, IDC_COMBO1, findCBox);
	DDX_Control(pDX, IDC_COMBO2, delCBox);
	DDX_Control(pDX, IDC_MF_COMBO, mfCBox);
	DDX_Text(pDX, IDC_MF_NUM, mfNum);
	DDX_Control(pDX, IDC_SALE_COMBO, saleCombo);
	DDX_Control(pDX, IDC_ADD_SN, addSNEdit);
	DDX_Control(pDX, IDC_MF_SN, mfSNEdit);
	DDX_Control(pDX, IDC_MF_NEWSN, newSnEdit);
	DDX_Control(pDX, IDC_TYPE_COMBO, typeCombo);
	DDX_Control(pDX, IDC_STAT_COMBO, statusCombo);
	DDX_Control(pDX, IDC_MF_STAT, mfStatusCombo);
	DDX_Control(pDX, IDC_MF_SALE, mfSaleCombo);
}

BEGIN_MESSAGE_MAP(CMBSNManagerDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_MESSAGE_VOID(WM_CLOSE, MyClose)
	ON_COMMAND(ID_ABOUT, CAboutDlg::OnAbout)
	ON_COMMAND(ID_OPEN, &CMBSNManagerDlg::OpenAccessData)
	ON_COMMAND(ID_CLOSE, &CMBSNManagerDlg::CloseAccessData)
	ON_COMMAND(ID_NEW, &CMBSNManagerDlg::CreateAccessData)
	ON_MESSAGE(WM_CREATEDB_ACCESS, OnCreateAccessChange)

	ON_BN_CLICKED(IDC_FIND_BTN, &CMBSNManagerDlg::OnBnClickedFindBtn)
	ON_BN_CLICKED(IDC_ADD_BTN, &CMBSNManagerDlg::OnBnClickedAddBtn)
	ON_COMMAND(ID_REF, &CMBSNManagerDlg::OnRef)
	ON_BN_CLICKED(IDC_DEL_BTN, &CMBSNManagerDlg::OnBnClickedDelBtn)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST1, OnItemchangedList1)
	ON_BN_CLICKED(IDC_MF_BTN, &CMBSNManagerDlg::OnBnClickedMfBtn)
	ON_CBN_SELCHANGE(IDC_MF_COMBO, &CMBSNManagerDlg::OnCbnSelchangeMfCombo)
END_MESSAGE_MAP()


// CMBSNManagerDlg 消息处理程序

BOOL CMBSNManagerDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	m_Menu.LoadMenu(IDR_MENU1);
	SetMenu(&m_Menu);

	OnInitView();
	OnInitDataBase();

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CMBSNManagerDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CMBSNManagerDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CMBSNManagerDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

CStringArray mbTypsArr;
CString accessFile;
CString accessPath;
CString accessName;
CString tableName;
CString errorMsg;
ADOTools ado;

CString status[2] = {
	_T("入库"),_T("出库")
};

void CMBSNManagerDlg::OnInitView()
{
	TCHAR filePath[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, filePath);
	SetCurrentDirectory(filePath);
	CString fileName;
	fileName.Format(_T("%s"), filePath);
	accessPath = fileName;

	CTime time;
	time = CTime::GetCurrentTime();
	CString dateStr = time.Format("%Y-%m-%d");
	GetDlgItem(IDC_ADD_DATE)->SetWindowText(dateStr);

	OnLoadMBTypes();

	statusCombo.InsertString(0, status[0]);
	statusCombo.InsertString(1, status[1]);
	mfStatusCombo.InsertString(0, status[0]);
	mfStatusCombo.InsertString(1, status[1]);

	findCBox.InsertString(0, _T("订单号"));
	findCBox.InsertString(1, _T("条码"));
	findCBox.InsertString(2, _T("状态"));
	findCBox.SetCurSel(0);
	delCBox.InsertString(0, _T("订单号"));
	delCBox.InsertString(1, _T("条码"));
	delCBox.SetCurSel(0);
	mfCBox.InsertString(0, _T("订单号"));
	mfCBox.InsertString(1, _T("条码"));
	mfCBox.SetCurSel(0);

	mList.InsertColumn(0, _T("NO"), LVCFMT_LEFT);
	mList.SetColumnWidth(0, 50);
	mList.InsertColumn(1, _T("订单号"), LVCFMT_LEFT, 120);
	mList.InsertColumn(2, _T("日期"), LVCFMT_LEFT, 80);
	mList.InsertColumn(3, _T("型号"), LVCFMT_LEFT, 130);
	mList.InsertColumn(4, _T("条码"), LVCFMT_LEFT, 130);
	mList.InsertColumn(5, _T("客户"), LVCFMT_LEFT, 70);
	mList.InsertColumn(6, _T("业务"), LVCFMT_LEFT, 60);
	mList.InsertColumn(7, _T("状态"), LVCFMT_LEFT, 60);
}

void CMBSNManagerDlg::OnLoadMBTypes()
{
	//CString types;
	//CString confifPath = accessPath + _T("\\config.ini");
	//GetPrivateProfileString(_T("MBTYPE"), _T("NAME"), _T(""), types.GetBuffer(MAX_PATH), MAX_PATH, confifPath);
	//CString szTemp;
	//int i = 0;
	//while (AfxExtractSubString(szTemp, types, i, ','))
	//{
	//	//mbTypsList.AddTail(szTemp);
	//	mbTypsArr.Add(szTemp);
	//	typeCombo.InsertString(i, szTemp);
	//	i++;
	//}
	////typeCombo.SetCurSel(0);

	CString modelPath = accessPath + _T("\\model.txt");
	CStdioFile file;
	if (!file.Open(modelPath, CFile::modeRead))
	{
		AfxMessageBox(_T("打开主板型号配置文件失败!"));
		return;
	}
	//CStringArray arr2;
	CString temp;
	int i = 0;
	while (file.ReadString(temp))
	{
		if (temp.Trim().GetLength()>0)
		{
			//arr2.Add(temp);
			typeCombo.InsertString(i, temp);
			i++;
		}
	}
	
	file.Close();
	/*CString str;
	for (int j=0;j<arr2.GetSize();j++)
	{
		str += arr2.GetAt(j)+_T(",");
	}
	AfxMessageBox(str);*/

	CString salePath = accessPath + _T("\\sales.txt");
	CStdioFile file2;
	if (!file2.Open(salePath, CFile::modeRead))
	{
		AfxMessageBox(_T("打开业务配置文件失败!"));
		return;
	}
	CString saleTemp;
	int j = 0;
	while (file2.ReadString(saleTemp))
	{
		if (saleTemp.Trim().GetLength() > 0)
		{
			saleCombo.InsertString(j, saleTemp);
			mfSaleCombo.InsertString(j, saleTemp);
			j++;
		}
	}
	file2.Close();
}

LRESULT CMBSNManagerDlg::OnCreateAccessChange(WPARAM wParam, LPARAM lParam)
{
	CloseAccessData();
	OnConDBAndUpdateList();
	return 0;
}

void CMBSNManagerDlg::OpenAccessData()
{
	TCHAR filePath[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, filePath);
	SetCurrentDirectory(filePath);
	
	// 设置过滤器   
	TCHAR szFilter[] = _T("数据库(*.accdb)|*.accdb");
	// 构造打开文件对话框   
	CFileDialog fileDlg(TRUE, _T("accdb"), NULL, 0, szFilter, this);

	// 显示打开文件对话框   
	if (IDOK == fileDlg.DoModal())
	{
		accessFile = fileDlg.GetPathName();
		accessPath = fileDlg.GetFolderPath();
		accessName = fileDlg.GetFileName();

		//AfxMessageBox(accessPath+_T("\r\n")+accessFile);
		CloseAccessData();
		OnConDBAndUpdateList();
	}
}

void CMBSNManagerDlg::CloseAccessData()
{
	ado.ExitADOConn();
	mList.DeleteAllItems();
	tableName.Empty();
}

BOOL CMBSNManagerDlg::OnInitDataFile()
{
	/*TCHAR filePath[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, filePath);
	SetCurrentDirectory(filePath);
	CString fileName;
	fileName.Format(_T("%s"), filePath);
	accessPath = fileName;*/

	CFileFind finder;
	BOOL isNotEmpty = finder.FindFile(accessPath + _T("\\*.accdb"));
	if (!isNotEmpty)
	{
		return FALSE;
	}
	else
	{
		finder.FindNextFile();
		accessName = finder.GetFileName();
		accessFile = accessPath + _T("\\") + accessName;
	}

	//AfxMessageBox(accessPath+_T("\r\n")+accessFile);

	return TRUE;
}

void CMBSNManagerDlg::OnInitDataBase()
{
	BOOL df = OnInitDataFile();
	if (!df)
	{
		GetDlgItem(IDC_DATA_TOTAL)->SetWindowText(_T("数据库不存在!"));
		return;
	}
	OnConDBAndUpdateList();
}

void CMBSNManagerDlg::OnConDBAndUpdateList()
{
	BOOL cret = ado.OnConnADODB();
	if (cret)
	{
		GetDlgItem(IDC_FIND_BTN)->EnableWindow(TRUE);
		GetDlgItem(IDC_ADD_BTN)->EnableWindow(TRUE);
		GetDlgItem(IDC_MF_BTN)->EnableWindow(TRUE);
		GetDlgItem(IDC_DEL_BTN)->EnableWindow(TRUE);

		CString lpSql;
		lpSql.Format(_T("SELECT * FROM %s"), tableName);
		vector<SNDATA> vecdata = ado.GetADODBForSql((LPCTSTR)lpSql);

		CString status;
		status.Format(_T("Total: %d	DB: %s	"), vecdata.size(), accessFile);
		GetDlgItem(IDC_DATA_TOTAL)->SetWindowText(status);

		mList.DeleteAllItems();
		int nItem = 0;
		for each (SNDATA data in vecdata)
		{
			CString id;
			//id.Format(_T("%d"), data.sid);
			id.Format(_T("%d"), nItem + 1);
			mList.InsertItem(nItem, id);
			mList.SetItemText(nItem, 1, data.order);
			mList.SetItemText(nItem, 2, data.ordate);
			mList.SetItemText(nItem, 3, data.model);
			mList.SetItemText(nItem, 4, data.sn);
			mList.SetItemText(nItem, 5, data.client);
			mList.SetItemText(nItem, 6, data.sale);
			mList.SetItemText(nItem, 7, data.status);

			nItem++;
		}
	}
}

void CMBSNManagerDlg::OnBnClickedFindBtn()
{
	CString findStr;
	TCHAR szSql[1024] = { 0 };

	GetDlgItemText(IDC_FIND, findStr);
	if (findStr.Trim().GetLength() == 0)
	{
		AfxMessageBox(_T("输入栏不能为空"));
	}
	if (findCBox.GetCurSel() == 0)
	{
		_stprintf(szSql, _T("SELECT * FROM %s WHERE OrderNo ='%s'"), (LPCTSTR)tableName, (LPCTSTR)findStr);
	}
	else if(findCBox.GetCurSel() == 1)
	{
		_stprintf(szSql, _T("SELECT * FROM %s WHERE SerialNo ='%s'"), (LPCTSTR)tableName, (LPCTSTR)findStr);
	}
	else
	{
		_stprintf(szSql, _T("SELECT * FROM %s WHERE Status ='%s'"), (LPCTSTR)tableName, (LPCTSTR)findStr);
	}

	vector<SNDATA> vecdata = ado.GetADODBForSql(szSql);
	CString total;
	total.Format(_T("Total: %d	DB: %s	"), vecdata.size(), accessFile);
	GetDlgItem(IDC_DATA_TOTAL)->SetWindowText(total);

	if (vecdata.size() > 0)
	{
		mList.DeleteAllItems();
		int nItem = 0;
		for each (SNDATA data in vecdata)
		{
			CString id;
			//id.Format(_T("%d"), data.sid);
			id.Format(_T("%d"), nItem + 1);
			mList.InsertItem(nItem, id);
			mList.SetItemText(nItem, 1, data.order);
			mList.SetItemText(nItem, 2, data.ordate);
			mList.SetItemText(nItem, 3, data.model);
			mList.SetItemText(nItem, 4, data.sn);
			mList.SetItemText(nItem, 5, data.client);
			mList.SetItemText(nItem, 6, data.sale);
			mList.SetItemText(nItem, 7, data.status);

			nItem++;
		}
	}
	else
	{
		AfxMessageBox(_T("没有找到数据"));
	}
}
BOOL IsNum(CString str)
{
	int n = str.GetLength();
	for (int i = 0; i < n; i++)
	{
		if ((str[i] < '0') || (str[i] > '9'))
			return FALSE;
	}
	return TRUE;
}

//新增数据
void CMBSNManagerDlg::OnBnClickedAddBtn()
{
	UpdateData(TRUE);
	int num = addNumEdit;
	CString order;
	CString date;
	CString model;
	CString client;
	CString sale;
	CString stat;
	CString sn;
	CString snPre;
	CString snSuf;
	//CString sn1;
	//CString sn2;
	CString temp;
	CString finalSN;

	GetDlgItemText(IDC_ADD_ORDER, order);
	GetDlgItemText(IDC_ADD_DATE, date);
	//GetDlgItemText(IDC_ADD_MODEL, model);
	//model = mbTypsArr.GetAt(typeCombo.GetCurSel());
	typeCombo.GetWindowText(model);
	GetDlgItemText(IDC_ADD_CLIENT, client);
	//GetDlgItemText(IDC_ADD_SALE, sale);
	//sale = sales[saleCombo.GetCurSel()];
	saleCombo.GetWindowText(sale);
	stat = status[statusCombo.GetCurSel()];
	//GetDlgItemText(IDC_ADD_SN1, sn1);
	//GetDlgItemText(IDC_ADD_SN2, sn2);
	addSNEdit.GetWindowText(sn);
	
	if (order.Trim().GetLength() == 0 || date.Trim().GetLength() == 0||
		model.Trim().GetLength() == 0 || sn.Trim().GetLength() == 0||
		sale.Trim().GetLength() == 0 || stat.Trim().GetLength() == 0)
	{
		AfxMessageBox(_T("数据不能为空"));
		return;
	}
	snPre = sn.Mid(0, 5);
	snSuf = sn.Mid(5, sn.Trim().GetLength());
	BOOL flag = IsNum(snSuf);
	if (!flag)
	{
		AfxMessageBox(_T("条码格式错误!"));
		return;
	}

	int length = snSuf.GetLength();
	LONGLONG snval = atoll((CT2A)snSuf);//将整数字符串转换未整数
	LONGLONG addNo;
	for (int i = 0; i < num; i++)
	{
		addNo = snval + i;
		temp.Format(_T("%0*lld"), length, addNo);
		finalSN = snPre + temp;

		TCHAR szSql[1024] = { 0 };
		_stprintf(szSql, _T("INSERT INTO %s(OrderNo,OrderDate,Model,SerialNo,Client,Sale,Status)\
							 VALUES('%s','%s','%s','%s','%s','%s','%s')"),
			(LPCTSTR)tableName,(LPCTSTR)order, (LPCTSTR)date, (LPCTSTR)model,
			(LPCTSTR)finalSN, (LPCTSTR)client, (LPCTSTR)sale, (LPCTSTR)stat);
		//_tprintf(szSql);

		BOOL ret = ado.OnExecuteADODB(szSql);
		if (!ret)
		{
			AfxMessageBox(errorMsg);
			return;
		}
	}

	RefListView();
}

void CMBSNManagerDlg::OnBnClickedMfBtn()
{
	UpdateData(TRUE);
	CString order;
	CString date;
	CString model;
	CString client;
	CString sale;
	CString stat;
	CString sn;
	CString newSn;
	//CString sn1;
	//CString sn2;
	CString snPre;
	CString snSuf;
	CString temp;
	CString finalSN;
	CString input;

	GetDlgItemText(IDC_MF_ORDER, order);
	GetDlgItemText(IDC_MF_DATE, date);
	GetDlgItemText(IDC_MF_MODEL, model);
	GetDlgItemText(IDC_MF_CLIENT, client);
	GetDlgItemText(IDC_MF_SALE, sale);
	GetDlgItemText(IDC_MF_STAT, stat);
	//GetDlgItemText(IDC_MF_SN1, sn1);
	//GetDlgItemText(IDC_MF_SN2, sn2);
	mfSNEdit.GetWindowText(sn);
	GetDlgItemText(IDC_MF_NEWSN, newSn);
	GetDlgItemText(IDC_MF_EDIT, input);

	snPre = newSn.Mid(0, 5);
	snSuf = newSn.Mid(5, newSn.Trim().GetLength());
	BOOL flag = IsNum(snSuf);
	if (!flag)
	{
		AfxMessageBox(_T("新条码格式错误!"));
		return;
	}
	TCHAR szSql[1024] = { 0 };

	//订单号批量修改数据
	if (mfCBox.GetCurSel() == 0)
	{
		int num = mfNum;
		//如果条码为空，那么SQL语句用订单号作为条件(不修改条码)
		if (num<=1 || sn.Trim().GetLength()==0)
		{
			_stprintf(szSql, _T("UPDATE %s SET OrderNo='%s',OrderDate='%s',Model='%s',\
						Client='%s',Sale='%s',Status='%s' WHERE OrderNo='%s'"),
				(LPCTSTR)tableName, (LPCTSTR)order, (LPCTSTR)date, (LPCTSTR)model,
				(LPCTSTR)client, (LPCTSTR)sale, (LPCTSTR)stat, (LPCTSTR)input);

			BOOL ret = ado.OnExecuteADODB(szSql);
			if (!ret)
			{
				AfxMessageBox(errorMsg);
				return;
			}
		}
		else
		{
			//批量修改的数据如果包括条码，那么SQL语句的条件需使用条码号
			//先查询出当前订单号的所有条码
			TCHAR snSql[1024] = { 0 };
			_stprintf(snSql, _T("SELECT SerialNo FROM %s WHERE OrderNo ='%s'"), (LPCTSTR)tableName, (LPCTSTR)input);
			vector<CString> snVec = ado.GetADODBSNForSql(snSql);
			int snsize = snVec.size();
			if (snsize <= 0)
			{
				AfxMessageBox(_T("没有找到条码,无法修改"));
				return;
			}
			
			int start = 0;
			
			for each (CString var in snVec)
			{
				if (lstrcmp(sn, var)==0)
				{
					break;
				}
				start++;
			}
			//修改的数量不能越界
			if (start+num > snsize-start)
			{
				num = snsize-start;
			}
			
			CString oldSN;
			int length = snSuf.GetLength();
			LONGLONG snval = atoll((CT2A)snSuf);
			LONGLONG mfNo;
			for (int i = 0; i < num; i++)
			{
				mfNo = snval + i;
				temp.Format(_T("%0*lld"), length, mfNo);
				finalSN = snPre + temp;
				oldSN = snVec[start+i];

				_stprintf(szSql, _T("UPDATE %s SET OrderNo='%s',OrderDate='%s',Model='%s',\
						Client='%s',Sale='%s',Status='%s',SerialNo='%s' WHERE SerialNo='%s'"),
							(LPCTSTR)tableName, (LPCTSTR)order, (LPCTSTR)date, (LPCTSTR)model,
							(LPCTSTR)client, (LPCTSTR)sale, (LPCTSTR)stat, (LPCTSTR)finalSN, (LPCTSTR)oldSN);

				BOOL ret = ado.OnExecuteADODB(szSql);
				if (!ret)
				{
					AfxMessageBox(errorMsg);
					return;
				}
			}
		}
		RefListView();
	}
	else
	{
		int num = mfNum;
		//条码修改单条数据
		if (num<=1)
		{
			input.Trim();
			if (newSn.Trim().GetLength() == 0)
			{
				newSn = input;
				SetDlgItemText(IDC_MF_NEWSN, newSn);
			}
			_stprintf(szSql, _T("UPDATE %s SET OrderNo='%s',OrderDate='%s',Model='%s',\
						Client='%s',Sale='%s',Status='%s', SerialNo='%s' WHERE SerialNo='%s'"),
				(LPCTSTR)tableName, (LPCTSTR)order, (LPCTSTR)date, (LPCTSTR)model,
				(LPCTSTR)client, (LPCTSTR)sale, (LPCTSTR)stat, (LPCTSTR)newSn, (LPCTSTR)input);
			BOOL ret = ado.OnExecuteADODB(szSql);
			if (!ret)
			{
				AfxMessageBox(errorMsg);
				return;
			}
		}
		else//批量修改条码
		{
			//先查询出当前订单号的所有条码
			TCHAR snSql[1024] = { 0 };
			_stprintf(snSql, _T("SELECT SerialNo FROM %s WHERE OrderNo ='%s'"), (LPCTSTR)tableName, (LPCTSTR)order);
			vector<CString> snVec = ado.GetADODBSNForSql(snSql);
			int snsize = snVec.size();
			if (snsize <= 0)
			{
				AfxMessageBox(_T("没有找到条码,无法修改"));
				return;
			}

			int start = 0;

			for each (CString var in snVec)
			{
				if (lstrcmp(sn, var) == 0)
				{
					break;
				}
				start++;
			}
			//修改的数量不能越界
			if (start + num > snsize - start)
			{
				num = snsize - start;
			}

			CString oldSN;
			int length = snSuf.GetLength();
			LONGLONG snval = atoll((CT2A)snSuf);
			LONGLONG mfNo;
			for (int i = 0; i < num; i++)
			{
				mfNo = snval + i;
				temp.Format(_T("%0*lld"), length, mfNo);
				finalSN = snPre + temp;
				oldSN = snVec[start + i];

				_stprintf(szSql, _T("UPDATE %s SET OrderNo='%s',OrderDate='%s',Model='%s',\
						Client='%s',Sale='%s',Status='%s',SerialNo='%s' WHERE SerialNo='%s'"),
					(LPCTSTR)tableName, (LPCTSTR)order, (LPCTSTR)date, (LPCTSTR)model,
					(LPCTSTR)client, (LPCTSTR)sale, (LPCTSTR)stat, (LPCTSTR)finalSN, (LPCTSTR)oldSN);

				BOOL ret = ado.OnExecuteADODB(szSql);
				if (!ret)
				{
					AfxMessageBox(errorMsg);
					return;
				}
			}
		}
		
		RefListView();
	}
}


void CMBSNManagerDlg::OnRef()
{
	RefListView();
}


void CMBSNManagerDlg::RefListView()
{
	CString lpSql;
	lpSql.Format(_T("SELECT * FROM %s"), tableName);
	
	vector<SNDATA> vecdata = ado.GetADODBForSql((LPCTSTR)lpSql);
	if (vecdata.size()==0)
	{
		return;
	}
	
	CString status;
	status.Format(_T("Total: %d	DB: %s	"), vecdata.size(), accessFile);
	GetDlgItem(IDC_DATA_TOTAL)->SetWindowText(status);

	mList.DeleteAllItems();
	int nItem = 0;
	for each (SNDATA data in vecdata)
	{
		CString id;
		//id.Format(_T("%d"), data.sid);
		id.Format(_T("%d"), nItem+1);
		mList.InsertItem(nItem, id);
		mList.SetItemText(nItem, 1, data.order);
		mList.SetItemText(nItem, 2, data.ordate);
		mList.SetItemText(nItem, 3, data.model);
		mList.SetItemText(nItem, 4, data.sn);
		mList.SetItemText(nItem, 5, data.client);
		mList.SetItemText(nItem, 6, data.sale);
		mList.SetItemText(nItem, 7, data.status);

		nItem++;
	}
	if (nItem > 0)
	{
		mList.EnsureVisible(nItem - 1, FALSE);
	}
}

void CMBSNManagerDlg::OnItemchangedList1(NMHDR* pNMHDR, LRESULT* pResult)
{
	POSITION pos = mList.GetFirstSelectedItemPosition();
	if (pos == NULL)
	{
		//没有任何行选中
	}
	else
	{
		while (pos)
		{
			int nItem = mList.GetNextSelectedItem(pos);
			CString orderstr = mList.GetItemText(nItem, 1);
			CString datestr = mList.GetItemText(nItem, 2);
			CString modelstr = mList.GetItemText(nItem, 3);
			CString snstr = mList.GetItemText(nItem, 4);
			CString clientstr = mList.GetItemText(nItem, 5);
			CString salestr = mList.GetItemText(nItem, 6);
			CString statstr = mList.GetItemText(nItem, 7);

			SetDlgItemText(IDC_MF_ORDER, orderstr);
			SetDlgItemText(IDC_MF_DATE, datestr);
			SetDlgItemText(IDC_MF_MODEL, modelstr);
			SetDlgItemText(IDC_MF_CLIENT, clientstr);
			SetDlgItemText(IDC_MF_SALE, salestr);
			SetDlgItemText(IDC_MF_STAT, statstr);
			mfSNEdit.SetWindowText(snstr);
			SetDlgItemText(IDC_MF_NEWSN, snstr);

			if (mfCBox.GetCurSel() == 0)
				SetDlgItemText(IDC_MF_EDIT, orderstr);
			else
				SetDlgItemText(IDC_MF_EDIT, snstr);

			if (delCBox.GetCurSel() == 0)
				SetDlgItemText(IDC_DEL_EDIT, orderstr);
			else
				SetDlgItemText(IDC_DEL_EDIT, snstr);
		}
	}
}

void CMBSNManagerDlg::OnBnClickedDelBtn()
{
	CString delStr;
	TCHAR szSql[1024] = { 0 };

	GetDlgItemText(IDC_DEL_EDIT, delStr);
	if (delStr.Trim().GetLength() == 0)
	{
		AfxMessageBox(_T("输入栏不能为空"));
	}
	if (delCBox.GetCurSel() == 0)
	{
		_stprintf(szSql, _T("DELETE FROM %s WHERE OrderNo='%s'"), (LPCTSTR)tableName, (LPCTSTR)delStr);
	}
	else
	{
		_stprintf(szSql, _T("DELETE FROM %s WHERE SerialNo='%s'"), (LPCTSTR)tableName, (LPCTSTR)delStr);
	}
	BOOL ret = ado.OnExecuteADODB(szSql);
	if (ret)
	{
		RefListView();
	}
}

void CMBSNManagerDlg::CreateAccessData()
{
	CreateData cData;
	cData.DoModal();
}

void CAboutDlg::OnAbout()
{
	CAboutDlg dlgAbout;
	dlgAbout.DoModal();
}

void CMBSNManagerDlg::MyClose()
{
	ado.ExitADOConn();
	
	this->OnClose();
}

void CMBSNManagerDlg::OnCbnSelchangeMfCombo()
{
	if (mfCBox.GetCurSel() == 0)
	{
		//GetDlgItem(IDC_MF_SN)->EnableWindow(TRUE);
	}
	else
	{
		//GetDlgItem(IDC_MF_SN)->EnableWindow(FALSE);
	}
}


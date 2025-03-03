
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
	DDX_Control(pDX, IDC_DEL_COMBO, delTypeCombo);
	DDX_Control(pDX, IDC_DATETIMEPICKER1, myDateTime);
}

BOOL CMBSNManagerDlg::PreTranslateMessage(MSG * pMsg)
{
	if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE) return TRUE;  //去掉esc退出功能	
	if (pMsg->message==WM_KEYDOWN && pMsg->wParam==VK_RETURN) return TRUE; //去掉回车退出功能，当然在这里也会导致整个窗体的失去回车效果	
	
	else  //其他正常
		return CDialogEx::PreTranslateMessage(pMsg);
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
	ON_BN_CLICKED(IDC_BUTTON1, &CMBSNManagerDlg::OnBnClickedButton1)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_DATETIMEPICKER1, &CMBSNManagerDlg::OnDtnDatetimechangeDatetimepicker1)
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
CString getTime;
ADOTools ado;
int NEWPOSITION = 0;

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

	myDateTime.SetFormat(_T("yyyy'-'MM'-'dd"));
	CTime tt;
	DWORD d = myDateTime.GetTime(tt);
	if (d == GDT_VALID)
	{
		getTime = tt.Format(_T("%Y-%m-%d"));
	}

	GetDlgItem(IDC_ADD_DATE)->SetWindowText(getTime);

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
	CString modelPath = accessPath + _T("\\model.txt");
	CStdioFile file;
	if (!file.Open(modelPath, CFile::modeRead))
	{
		AfxMessageBox(_T("打开主板型号配置文件失败!"));
		return;
	}
	CString temp;
	int i = 0;
	while (file.ReadString(temp))
	{
		if (temp.Trim().GetLength()>0)
		{
			//arr2.Add(temp);
			typeCombo.InsertString(i, temp);
			delTypeCombo.InsertString(i, temp);
			i++;
		}
	}
	
	file.Close();

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

//初始化数据库，获取数据并填充到ListView，只获取当前月份
void CMBSNManagerDlg::OnConDBAndUpdateList()
{
	BOOL cret = ado.OnConnADODB();
	if (cret)
	{
		GetDlgItem(IDC_FIND_BTN)->EnableWindow(TRUE);
		GetDlgItem(IDC_ADD_BTN)->EnableWindow(TRUE);
		GetDlgItem(IDC_MF_BTN)->EnableWindow(TRUE);
		GetDlgItem(IDC_DEL_BTN)->EnableWindow(TRUE);

		CString szDate;
		szDate = getTime.Left(7);

		CString lpSql;
		//lpSql.Format(_T("SELECT * FROM %s "), tableName);
		lpSql.Format(_T("SELECT * FROM %s WHERE OrderDate like '%s-[0-9][0-9]' ORDER BY ID"), tableName, szDate);
		//lpSql.Format(_T("SELECT * FROM %s ORDER BY ID"), tableName);
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
		

		if (nItem > 0)
		{
			mList.EnsureVisible(nItem - 1, FALSE);
		}
	}
}

//查询订单数据
void CMBSNManagerDlg::OnBnClickedFindBtn()
{
	CString findStr;
	CString typeSN;
	TCHAR szSql[1024] = { 0 };

	GetDlgItemText(IDC_FIND, findStr);

	if (findStr.Trim().GetLength() == 0)
	{
		AfxMessageBox(_T("输入栏不能为空"));
	}
	//通过订单号查询
	if (findCBox.GetCurSel() == 0)
	{
		_stprintf(szSql, _T("SELECT * FROM %s WHERE OrderNo ='%s'"), (LPCTSTR)tableName, (LPCTSTR)findStr);
	}
	else if (findCBox.GetCurSel() == 1)
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
	CString typeSN;

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

	//只有批量添加数据才限制条码格式(有可能会有奇怪的格式)
	if (num > 1)
	{
		snPre = sn.Mid(0, 5);
		snSuf = sn.Mid(5, sn.Trim().GetLength());
		BOOL flag = IsNum(snSuf);//前5位不限，后面的数据必须为数字
		if (!flag)
		{
			AfxMessageBox(_T("条码格式错误!"));
			return;
		}

		int length = snSuf.GetLength();
		LONGLONG snval = atoll((CT2A)snSuf);//将整数字符串转换为整数
		LONGLONG addNo;

		int len = mList.GetItemCount();
		for (int i = 0; i < num; i++)
		{
			addNo = snval + i;
			temp.Format(_T("%0*lld"), length, addNo);
			finalSN = snPre + temp;
			typeSN = model + _T("_") + finalSN;

			TCHAR szSql[1024] = { 0 };
			_stprintf(szSql, _T("INSERT INTO %s(OrderNo,OrderDate,Model,SerialNo,Client,Sale,Status,TYPE_SN)\
							 VALUES('%s','%s','%s','%s','%s','%s','%s','%s')"),
				(LPCTSTR)tableName, (LPCTSTR)order, (LPCTSTR)date, (LPCTSTR)model,
				(LPCTSTR)finalSN, (LPCTSTR)client, (LPCTSTR)sale, (LPCTSTR)stat, (LPCTSTR)typeSN);
			//_tprintf(szSql);

			BOOL ret = ado.OnExecuteADODB(szSql);
			if (!ret)
			{
				AfxMessageBox(errorMsg);
				//重新连接数据库，否则后续添加的数据会失败
				ado.ExitADOConn();
				ado.OnConnADODB();
				return;
			}
			
			//刷新ListView
			CString id;
			id.Format(_T("%d"), len + i + 1);
			mList.InsertItem(len + i, id);
			mList.SetItemText(len + i, 1, order);
			mList.SetItemText(len + i, 2, date);
			mList.SetItemText(len + i, 3, model);
			mList.SetItemText(len + i, 4, finalSN);
			mList.SetItemText(len + i, 5, client);
			mList.SetItemText(len + i, 6, sale);
			mList.SetItemText(len + i, 7, stat);
		}
		mList.EnsureVisible(len + num -1, FALSE);
	}
	else
	{
		typeSN = model + _T("_") + sn;
		TCHAR szSql[1024] = { 0 };
		_stprintf(szSql, _T("INSERT INTO %s(OrderNo,OrderDate,Model,SerialNo,Client,Sale,Status,TYPE_SN)\
							 VALUES('%s','%s','%s','%s','%s','%s','%s','%s')"),
			(LPCTSTR)tableName, (LPCTSTR)order, (LPCTSTR)date, (LPCTSTR)model,
			(LPCTSTR)sn, (LPCTSTR)client, (LPCTSTR)sale, (LPCTSTR)stat, (LPCTSTR)typeSN);
		//_tprintf(szSql);

		BOOL ret = ado.OnExecuteADODB(szSql);
		if (!ret)
		{
			AfxMessageBox(errorMsg);
			ado.ExitADOConn();
			ado.OnConnADODB();
			return;
		}
		int len = mList.GetItemCount();
		CString id;
		id.Format(_T("%d"), len + 1);
		mList.InsertItem(len, id);
		mList.SetItemText(len, 1, order);
		mList.SetItemText(len, 2, date);
		mList.SetItemText(len, 3, model);
		mList.SetItemText(len, 4, sn);
		mList.SetItemText(len, 5, client);
		mList.SetItemText(len, 6, sale);
		mList.SetItemText(len, 7, stat);
	}
	//RefListView(0);

	CString total;
	total.Format(_T("Total: %d	DB: %s  "), mList.GetItemCount(), accessFile);
	GetDlgItem(IDC_DATA_TOTAL)->SetWindowText(total);
}

//修改数据
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
	CString typeSN;

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

	TCHAR szSql[1024] = { 0 };

	int mfCBoxIndex = mfCBox.GetCurSel();
	//订单号批量修改数据
	if (mfCBoxIndex == 0)
	{
		int num = mfNum;
		//如果条码为空，那么SQL语句用订单号作为条件(不修改条码)
		//num:数量 sn:当前条码
		if (num <= 1 || sn.Trim().GetLength() == 0)
		{
			if (model.Trim().GetLength()==0)
			{
				AfxMessageBox(_T("型号不能为空"));
				return;
			}
			//判断是否需要更新唯一字段TYPE_SN
			//先查询出订单号的主板型号
			TCHAR mySql[1024] = { 0 };
			_stprintf(mySql, _T("SELECT * FROM %s WHERE OrderNo ='%s'"), (LPCTSTR)tableName, (LPCTSTR)input);
			vector<SNDATA> vecdata = ado.GetADODBForSql(mySql);
			if (vecdata.size() <= 0)
			{
				AfxMessageBox(_T("没有找到数据，无法修改!"));
				return;
			}
			else
			{
				SNDATA data = vecdata[0];
				//如果型号相同，那么不需要修改型号，条码，TYPE_SN三个字段
				if (lstrcmp(data.model, model) == 0)
				{
					_stprintf(szSql, _T("UPDATE %s SET OrderNo='%s',OrderDate='%s',\
						Client='%s',Sale='%s',Status='%s' WHERE OrderNo='%s'"),
						(LPCTSTR)tableName, (LPCTSTR)order, (LPCTSTR)date, (LPCTSTR)client,
						(LPCTSTR)sale, (LPCTSTR)stat, (LPCTSTR)input);

					BOOL ret = ado.OnExecuteADODB(szSql);
					if (!ret)
					{
						AfxMessageBox(errorMsg);
						return;
					}
					//更新List
					CString str;
					for (int i = 0; i < mList.GetItemCount(); i++)
					{
						str = mList.GetItemText(i, 1);
						if (lstrcmp(str, input) == 0)
						{
							mList.SetItemText(i, 1, order);
							mList.SetItemText(i, 2, date);
							mList.SetItemText(i, 5, client);
							mList.SetItemText(i, 6, sale);
							mList.SetItemText(i, 7, stat);
						}
					}
				}
				//型号有变动，不需要修改条码，需要更新唯一字段TYPE_SN
				else
				{
					CString oldTypeSN;
					for each (SNDATA data in vecdata)
					{
						//更新唯一字段
						typeSN = model + _T("_") + data.sn;
						oldTypeSN = data.model + _T("_") + data.sn;
						_stprintf(szSql, _T("UPDATE %s SET OrderNo='%s',OrderDate='%s',Model='%s',\
						Client='%s',Sale='%s',Status='%s',TYPE_SN='%s' WHERE TYPE_SN='%s'"),
							(LPCTSTR)tableName, (LPCTSTR)order, (LPCTSTR)date, (LPCTSTR)model,
							(LPCTSTR)client, (LPCTSTR)sale, (LPCTSTR)stat, (LPCTSTR)typeSN, (LPCTSTR)oldTypeSN);

						BOOL ret = ado.OnExecuteADODB(szSql);
						if (!ret)
						{
							AfxMessageBox(errorMsg);
							return;
						}
					}
					//更新List
					CString str;
					for (int i = 0; i < mList.GetItemCount(); i++)
					{
						str = mList.GetItemText(i, 1);
						if (lstrcmp(str, input) == 0)
						{
							mList.SetItemText(i, 1, order);
							mList.SetItemText(i, 2, date);
							mList.SetItemText(i, 3, model);
							mList.SetItemText(i, 5, client);
							mList.SetItemText(i, 6, sale);
							mList.SetItemText(i, 7, stat);
						}
					}
				}
			}
		}
		//条码有变化
		else
		{	
			snPre = newSn.Mid(0, 5);
			snSuf = newSn.Mid(5, newSn.Trim().GetLength());
			BOOL flag = IsNum(snSuf);
			if (!flag)
			{
				AfxMessageBox(_T("新条码格式错误!"));
				return;
			}

			//批量修改的数据如果包括条码，那么SQL语句的条件需使用TYPE_SN
			//先查询出当前订单号的TYPE_SN
			TCHAR snSql[1024] = { 0 };
			CString filed = _T("TYPE_SN");
			_stprintf(snSql, _T("SELECT %s FROM %s WHERE OrderNo ='%s'"), (LPCTSTR)filed, (LPCTSTR)tableName, (LPCTSTR)input);
			vector<CString> oldTypeSNVec = ado.GetMyADODBForSql(snSql, (LPCTSTR)filed);
			int snsize = oldTypeSNVec.size();
			if (snsize <= 0)
			{
				AfxMessageBox(_T("没有找到条码,无法修改"));
				return;
			}
			
			int start = 0; //计算起始条码位置
			for each (CString var in oldTypeSNVec)
			{
				if (var.Find(sn,0) > -1)
				{
					break;
				}
				start++;
			}
			//修改的数量不能超出可修改的条码数
			if (num > snsize-start)
			{
				num = snsize-start;
			}
			
			int updateIndex = 0;
			for (int i=0; i<mList.GetItemCount();i++)
			{
				if (lstrcmp(sn,mList.GetItemText(i,4))==0)
				{
					updateIndex = i;
				}
			}

			//CString oldSN;
			CString oldTypeSN;
			int length = snSuf.GetLength();
			LONGLONG snval = atoll((CT2A)snSuf);
			LONGLONG mfNo;
			for (int i = 0; i < num; i++)
			{
				mfNo = snval + i;
				temp.Format(_T("%0*lld"), length, mfNo);
				finalSN = snPre + temp;//格式化后最终的条码
				typeSN = model + _T("_") + finalSN;//更新唯一字段TYPE_SN
				//oldSN = snVec[start+i];
				//oldTypeSN = model + _T("_") + oldSN;//where条件，当前唯一字段
				oldTypeSN = oldTypeSNVec[start + i];

				_stprintf(szSql, _T("UPDATE %s SET OrderNo='%s',OrderDate='%s',Model='%s',\
						Client='%s',Sale='%s',Status='%s',SerialNo='%s',TYPE_SN='%s' WHERE TYPE_SN='%s'"),
							(LPCTSTR)tableName, (LPCTSTR)order, (LPCTSTR)date, (LPCTSTR)model,
							(LPCTSTR)client, (LPCTSTR)sale, (LPCTSTR)stat, (LPCTSTR)finalSN,
							(LPCTSTR)typeSN, (LPCTSTR)oldTypeSN);

				BOOL ret = ado.OnExecuteADODB(szSql);
				if (!ret)
				{
					AfxMessageBox(errorMsg);
					return;
				}

				//更新List数据
				mList.SetItemText(updateIndex + i, 1, order);
				mList.SetItemText(updateIndex + i, 2, date);
				mList.SetItemText(updateIndex + i, 3, model);
				mList.SetItemText(updateIndex + i, 4, finalSN);
				mList.SetItemText(updateIndex + i, 5, client);
				mList.SetItemText(updateIndex + i, 6, sale);
				mList.SetItemText(updateIndex + i, 7, stat);
			}
		}
	}
	//通过条码修改数据
	else
	{
		int num = mfNum;
		//修改单条数据(不限制条码格式)
		if (num <= 1)
		{
			//需先获取该条码的型号
			TCHAR mySql[1024] = { 0 };
			_stprintf(mySql, _T("SELECT TYPE_SN FROM %s WHERE SerialNo ='%s'"), (LPCTSTR)tableName, (LPCTSTR)input);
			CString filed = _T("TYPE_SN");
			CString oldTypeSN;
			oldTypeSN = ado.GetADODBValueForSql(mySql, (LPCTSTR)filed);
			if (oldTypeSN.Trim().GetLength() == 0)
			{
				AfxMessageBox(_T("无法修改型号!"));
				return;
			}

			input.Trim();
			//CString cTypeSN = model + _T("_") + input;
			if (newSn.Trim().GetLength() == 0)
			{
				newSn = input;
				SetDlgItemText(IDC_MF_NEWSN, newSn);
			}
			typeSN = model + _T("_") + newSn;//更新唯一TYPE_SN
			_stprintf(szSql, _T("UPDATE %s SET OrderNo='%s',OrderDate='%s',Model='%s',\
						Client='%s',Sale='%s',Status='%s', SerialNo='%s', TYPE_SN='%s' WHERE TYPE_SN='%s'"),
				(LPCTSTR)tableName, (LPCTSTR)order, (LPCTSTR)date, (LPCTSTR)model,
				(LPCTSTR)client, (LPCTSTR)sale, (LPCTSTR)stat, (LPCTSTR)newSn, (LPCTSTR)typeSN, (LPCTSTR)oldTypeSN);
			BOOL ret = ado.OnExecuteADODB(szSql);
			if (!ret)
			{
				AfxMessageBox(errorMsg);
				return;
			}
			//更新List
			int index = 0;
			CString str;
			for (int i = 0; i < mList.GetItemCount(); i++)
			{
				str = mList.GetItemText(i, 4);
				if (lstrcmp(str, input)==0)
				{
					index = i;
					break;
				}
			}
			mList.SetItemText(index, 1, order);
			mList.SetItemText(index, 2, date);
			mList.SetItemText(index, 3, model);
			mList.SetItemText(index, 4, newSn);
			mList.SetItemText(index, 5, client);
			mList.SetItemText(index, 6, sale);
			mList.SetItemText(index, 7, stat);
		}
		//批量修改条码信息
		else
		{	
			snPre = newSn.Mid(0, 5);
			snSuf = newSn.Mid(5, newSn.Trim().GetLength());
			BOOL flag = IsNum(snSuf);
			if (!flag)
			{
				AfxMessageBox(_T("新条码格式错误!"));
				return;
			}
			//1.先查询除当前条码的订单号
			TCHAR orSql[1024] = { 0 };
			_stprintf(orSql, _T("SELECT OrderNo FROM %s WHERE SerialNo ='%s'"), (LPCTSTR)tableName, (LPCTSTR)input);
			CString ONo = _T("OrderNo");
			CString oldOrderNo;
			oldOrderNo = ado.GetADODBValueForSql(orSql, (LPCTSTR)ONo);
			if (oldOrderNo.Trim().GetLength() == 0)
			{
				AfxMessageBox(_T("无法修改订单号!"));
				return;
			}

			//2.再查询出当前订单号的所有TYPE_SN
			TCHAR snSql[1024] = { 0 };
			CString filed = _T("TYPE_SN");
			_stprintf(snSql, _T("SELECT %s FROM %s WHERE OrderNo ='%s'"), (LPCTSTR)filed, (LPCTSTR)tableName, (LPCTSTR)oldOrderNo);
			vector<CString> oldTypeSNVec = ado.GetMyADODBForSql(snSql, (LPCTSTR)filed);
			int snsize = oldTypeSNVec.size();
			if (snsize <= 0)
			{
				AfxMessageBox(_T("没有找到条码,无法修改"));
				return;
			}

			int start = 0;
			for each (CString var in oldTypeSNVec)
			{
				if (var.Find(sn, 0) > -1)
				{
					break;
				}
				start++;
			}
			int updateIndex = 0;
			for (int i = 0; i < mList.GetItemCount(); i++)
			{
				if (lstrcmp(sn, mList.GetItemText(i, 4)) == 0)
				{
					updateIndex = i;
				}
			}
			
			//修改的数量不能越界
			if (num > snsize - start)
			{
				num = snsize - start;
			}

			//CString oldSN;
			CString oldTypeSN;
			int length = snSuf.GetLength();
			LONGLONG snval = atoll((CT2A)snSuf);
			LONGLONG mfNo;
			for (int i = 0; i < num; i++)
			{
				mfNo = snval + i;
				temp.Format(_T("%0*lld"), length, mfNo);
				finalSN = snPre + temp;
				//oldSN = snVec[start + i];
				typeSN = model + _T("_") + finalSN;//更新唯一字段TYPE_SN
				//oldTypeSN = model + _T("_") + oldSN;//where条件，当前唯一字段
				oldTypeSN = oldTypeSNVec[start + i];

				_stprintf(szSql, _T("UPDATE %s SET OrderNo='%s',OrderDate='%s',Model='%s',\
						Client='%s',Sale='%s',Status='%s',SerialNo='%s',TYPE_SN='%s' WHERE TYPE_SN='%s'"),
					(LPCTSTR)tableName, (LPCTSTR)order, (LPCTSTR)date, (LPCTSTR)model,
					(LPCTSTR)client, (LPCTSTR)sale, (LPCTSTR)stat, (LPCTSTR)finalSN, (LPCTSTR)typeSN, (LPCTSTR)oldTypeSN);

				BOOL ret = ado.OnExecuteADODB(szSql);
				if (!ret)
				{
					AfxMessageBox(errorMsg);
					return;
				}
				//更新List数据
				mList.SetItemText(updateIndex + i, 1, order);
				mList.SetItemText(updateIndex + i, 2, date);
				mList.SetItemText(updateIndex + i, 3, model);
				mList.SetItemText(updateIndex + i, 4, finalSN);
				mList.SetItemText(updateIndex + i, 5, client);
				mList.SetItemText(updateIndex + i, 6, sale);
				mList.SetItemText(updateIndex + i, 7, stat);
			}
		}

		//RefListView(1);
	}
}



//删除数据
void CMBSNManagerDlg::OnBnClickedDelBtn()
{
	CString delStr;
	CString typeStr;
	CString typeSN;
	TCHAR szSql[1024] = { 0 };

	GetDlgItemText(IDC_DEL_EDIT, delStr);
	delTypeCombo.GetWindowText(typeStr);
	if (delStr.Trim().GetLength() == 0)
	{
		AfxMessageBox(_T("输入栏不能为空"));
	}
	//通过订单号删除
	if (delCBox.GetCurSel() == 0)
	{
		//如果型号为空，那么订单号作为唯一条件
		if (typeStr.Trim().GetLength() == 0)
		{
			_stprintf(szSql, _T("DELETE FROM %s WHERE OrderNo='%s'"), (LPCTSTR)tableName, (LPCTSTR)delStr);
		}
		else
		{
			_stprintf(szSql, _T("DELETE FROM %s WHERE OrderNo='%s' AND Model='%s'"), (LPCTSTR)tableName, (LPCTSTR)delStr, (LPCTSTR)typeStr);
		}
		
		BOOL ret = ado.OnExecuteADODB(szSql);
		if (ret)
		{
			int len = mList.GetItemCount();
			int index = 0;
			CString str;
			for (int i = len -1; i >= 0; i--)
			{
				str = mList.GetItemText(i, 1);
				if (lstrcmp(str, delStr)==0)
				{
					//arr.push_back(i);
					mList.DeleteItem(i);
				}
			}

			//刷新List元素ID
			for (int j=0;j<mList.GetItemCount();j++)
			{
				CString id;
				id.Format(_T("%d"), j + 1);
				mList.SetItemText(j, 0, id);
			}

			CString total;
			total.Format(_T("Total: %d	DB: %s	"), mList.GetItemCount(), accessFile);
			GetDlgItem(IDC_DATA_TOTAL)->SetWindowText(total);
		}
	}
	else
	{
		//通过序列号删除，会删除重复序列号的数据
		if (typeStr.Trim().GetLength()==0)
		{
			_stprintf(szSql, _T("DELETE FROM %s WHERE SerialNo='%s'"), (LPCTSTR)tableName, (LPCTSTR)delStr);
		}
		else//通过TYPE_SN删除
		{
			typeSN = typeStr + _T("_") + delStr;
			_stprintf(szSql, _T("DELETE FROM %s WHERE TYPE_SN='%s'"), (LPCTSTR)tableName, (LPCTSTR)typeSN);
		}
		BOOL ret = ado.OnExecuteADODB(szSql);
		if (ret)
		{
			//不查询数据库更新ListView，删除不要的数据
			int index = 0;
			CString str;
			for (int i = 0; i< mList.GetItemCount();i++)
			{
				str = mList.GetItemText(i, 4);
				if (!str.Compare(delStr))
				{
					index = i;
					break;
				}
			}
			mList.DeleteItem(index);
			//刷新List元素ID
			for (int j = 0; j < mList.GetItemCount(); j++)
			{
				CString id;
				id.Format(_T("%d"), j + 1);
				mList.SetItemText(j, 0, id);
			}

			CString total;
			total.Format(_T("Total: %d	DB: %s	"), mList.GetItemCount(), accessFile);
			GetDlgItem(IDC_DATA_TOTAL)->SetWindowText(total);
		}
	}
}

//刷新数据
void CMBSNManagerDlg::OnRef()
{
	RefListView(0);
}

//查询数据刷新ListView 0:滚动到最后一行 1:当前行
void CMBSNManagerDlg::RefListView(int pos)
{
	CString lpSql;
	CString szDate;
	szDate = getTime.Left(7);
	lpSql.Format(_T("SELECT * FROM %s WHERE OrderDate like '%s-[0-9][0-9]' ORDER BY ID"), tableName, szDate);

	vector<SNDATA> vecdata = ado.GetADODBForSql((LPCTSTR)lpSql);
	if (vecdata.size() == 0)
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
	if (nItem > 0)
	{
		if (pos == 0)
		{
			mList.EnsureVisible(nItem - 1, FALSE);
		}
		else
		{
			mList.EnsureVisible(NEWPOSITION, FALSE);
		}
	}
}

//监听ListView点击事件
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
			NEWPOSITION = nItem;
			CString orderstr = mList.GetItemText(nItem, 1);
			CString datestr = mList.GetItemText(nItem, 2);
			CString modelstr = mList.GetItemText(nItem, 3);
			CString snstr = mList.GetItemText(nItem, 4);
			CString clientstr = mList.GetItemText(nItem, 5);
			CString salestr = mList.GetItemText(nItem, 6);
			CString statstr = mList.GetItemText(nItem, 7);

			SetDlgItemText(IDC_FIND, orderstr);
			SetDlgItemText(IDC_MF_ORDER, orderstr);
			SetDlgItemText(IDC_MF_DATE, datestr);
			SetDlgItemText(IDC_MF_MODEL, modelstr);
			SetDlgItemText(IDC_MF_CLIENT, clientstr);
			SetDlgItemText(IDC_MF_SALE, salestr);
			SetDlgItemText(IDC_MF_STAT, statstr);
			mfSNEdit.SetWindowText(snstr);
			SetDlgItemText(IDC_MF_NEWSN, snstr);

			if (findCBox.GetCurSel() == 0)
			{
				SetDlgItemText(IDC_FIND, orderstr);
			}
			else if (findCBox.GetCurSel() == 1)
			{
				SetDlgItemText(IDC_FIND, snstr);
			}
			else
			{
				SetDlgItemText(IDC_FIND, statstr);
			}

			if (mfCBox.GetCurSel() == 0)
			{
				SetDlgItemText(IDC_MF_EDIT, orderstr);
			}
			else
			{
				SetDlgItemText(IDC_MF_EDIT, snstr);
			}

			if (delCBox.GetCurSel() == 0)
			{
				SetDlgItemText(IDC_DEL_EDIT, orderstr);
			}
			else
			{
				SetDlgItemText(IDC_DEL_EDIT, snstr);
			}
			SetDlgItemText(IDC_DEL_COMBO, modelstr);
		}
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

//克隆到新表(已废弃)
void CMBSNManagerDlg::OnBnClickedButton1()
{
	TCHAR filePath[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, filePath);
	SetCurrentDirectory(filePath);

	//2.获取所有的数据
	CString lpSql;
	lpSql.Format(_T("SELECT * FROM %s"), tableName);
	vector<SNDATA> vecdata = ado.GetADODBForSql((LPCTSTR)lpSql);
	
	if (vecdata.size() > 0)
	{
		//3.打开新数据库
		CloseAccessData();

		//1.创建一张新的表，名字为newnewnew
		CString dbName = accessPath + _T("\\newnewnew.accdb");
		accessFile = dbName;

		BOOL ret = ado.CreateADOData(dbName);
		if (!ret)
		{
			AfxMessageBox(_T("数据库创建失败!"));
			return;
		}


		OnConDBAndUpdateList();

		//4.添加数据
		TCHAR szSql[1024] = { 0 };
		CString typeSN;
		for each (SNDATA data in vecdata)
		{
			typeSN.Format(_T("%s_%s"), (LPCTSTR)data.model, (LPCTSTR)data.sn);

			_stprintf(szSql, _T("INSERT INTO %s(OrderNo,OrderDate,Model,SerialNo,Client,Sale,Status,TYPE_SN)\
							 VALUES('%s','%s','%s','%s','%s','%s','%s','%s')"),
				(LPCTSTR)tableName, (LPCTSTR)data.order, (LPCTSTR)data.ordate, (LPCTSTR)data.model,
				(LPCTSTR)data.sn, (LPCTSTR)data.client, (LPCTSTR)data.sale, (LPCTSTR)data.status, (LPCTSTR)typeSN);

			BOOL ret = ado.OnExecuteADODB(szSql);
			if (!ret)
			{
				AfxMessageBox(errorMsg);
				//重新连接数据库，否则后续添加的数据会失败
				ado.ExitADOConn();
				ado.OnConnADODB();
				return;
			}

		}
		RefListView(1);
	}

	

}

void CMBSNManagerDlg::OnDtnDatetimechangeDatetimepicker1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMDATETIMECHANGE pDTChange = reinterpret_cast<LPNMDATETIMECHANGE>(pNMHDR);
	// TODO: 在此添加控件通知处理程序代码
	if (pDTChange->dwFlags == GDT_NONE)
	{
		myDateTime.EnableWindow(FALSE);
	}
	else if(pDTChange->dwFlags == GDT_VALID)
	{
		if (myDateTime.IsWindowEnabled() == TRUE)
		{
			CTime tt;
			DWORD d = myDateTime.GetTime(tt);
			if (d == GDT_VALID)
			{
				getTime = tt.Format(_T("%Y-%m-%d"));
				OnRef();
			}
		}
	}
	*pResult = 0;
}

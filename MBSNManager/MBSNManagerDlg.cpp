
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
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMBSNManagerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, mList);
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
	ON_MESSAGE(WM_ONINIT_ACCESS, OnInitAccessChange)

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

	OnInitDBList();
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

void CMBSNManagerDlg::OnInitDBList()
{
	mList.InsertColumn(0, _T("ID"), LVCFMT_LEFT);
	mList.SetColumnWidth(0, 50);
	mList.InsertColumn(1, _T("订单号"), LVCFMT_LEFT, 100);
	mList.InsertColumn(2, _T("日期"), LVCFMT_LEFT, 100);
	mList.InsertColumn(3, _T("型号"), LVCFMT_LEFT, 100);
	mList.InsertColumn(4, _T("条码"), LVCFMT_LEFT, 100);
	mList.InsertColumn(5, _T("客户"), LVCFMT_LEFT, 50);
	mList.InsertColumn(6, _T("业务"), LVCFMT_LEFT, 50);
}

CString accessPath;
CString accessName;
ADOTools ado;

void CMBSNManagerDlg::OnInitDataBase()
{
	//AfxMessageBox(accessPath + _T(" ") + accessName);
	TCHAR filePath[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, filePath);

	CString fileName;
	fileName.Format(_T("%s"), filePath);
	fileName = fileName + _T("\\test.accdb");

	if (!PathFileExists(fileName))
	{
		AfxMessageBox(_T("数据库不存在!"));
		return;
	}

	BOOL cret = ado.OnConnADODB();
	if (cret)
	{

	}
	

}

LRESULT CMBSNManagerDlg::OnInitAccessChange(WPARAM wParam, LPARAM lParam)
{
	//OnInitDataBase();
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
		accessPath = fileDlg.GetPathName();
		accessName = fileDlg.GetFileName()+_T(".")+fileDlg.GetFileExt();
	}

}

void CMBSNManagerDlg::CloseAccessData()
{

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
	
	this->OnClose();
}
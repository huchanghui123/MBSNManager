﻿// CreateData.cpp: 实现文件
//

#include "pch.h"
#include "MBSNManager.h"
#include "CreateData.h"
#include "afxdialogex.h"
#include "ADOTools.h"

// CreateData 对话框

IMPLEMENT_DYNAMIC(CreateData, CDialogEx)

CreateData::CreateData(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DIALOG1, pParent)
{

}

CreateData::~CreateData()
{
}

void CreateData::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CreateData, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON1, &CreateData::OnBnClickedButton1)
END_MESSAGE_MAP()


// CreateData 消息处理程序

void CreateData::OnBnClickedButton1()
{
	CString dbName;
	GetDlgItemText(IDC_EDIT1, dbName);

	if (dbName.Trim().GetLength() == 0)
	{
		return;
	}
	dbName = accessPath + _T("\\") + dbName + _T(".accdb");
	ADOTools ado;
	BOOL ret = ado.CreateADOData(dbName);
	if (ret)
	{
		AfxMessageBox(_T("数据库创建成功!"));
		accessFile = dbName;
		GetParent()->SendMessage(WM_CREATEDB_ACCESS, 0, 0);
		SendMessage(WM_CLOSE);
	}

}



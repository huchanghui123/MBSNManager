﻿#pragma once


// CreateData 对话框

class CreateData : public CDialogEx
{
	DECLARE_DYNAMIC(CreateData)

public:
	CreateData(CWnd* pParent = nullptr);   // 标准构造函数
	virtual ~CreateData();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG1 };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
};

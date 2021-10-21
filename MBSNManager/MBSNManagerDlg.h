
// MBSNManagerDlg.h: 头文件
//

#pragma once
#include "manager.h"

// CMBSNManagerDlg 对话框
class CMBSNManagerDlg : public CDialogEx
{
// 构造
public:
	CMBSNManagerDlg(CWnd* pParent = nullptr);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_MBSNMANAGER_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持

// 实现
protected:
	HICON m_hIcon;
	CMenu m_Menu;

	// 生成的消息映射函数
	afx_msg void OpenAccessData();
	afx_msg void CloseAccessData();
	afx_msg void CreateAccessData();

	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg void MyClose();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

	void OnInitDataBase();
	afx_msg  LRESULT  OnInitAccessChange(WPARAM wParam, LPARAM lParam);

public:
	

};

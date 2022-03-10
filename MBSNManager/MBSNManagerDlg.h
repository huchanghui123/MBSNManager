
// MBSNManagerDlg.h: 头文件
//

#pragma once
#include "manager.h"
#include "MyCEditEx.h"
#include "CAutoCombox.h"

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

	void OnInitView();
	void OnLoadMBTypes();
	BOOL OnInitDataFile();
	void OnInitDataBase();
	void OnConDBAndUpdateList();
	void RefListView(int);
	afx_msg  LRESULT  OnCreateAccessChange(WPARAM wParam, LPARAM lParam);

public:
	CListCtrl mList;
	int addNumEdit;
	afx_msg void OnRef();
	afx_msg void OnBnClickedFindBtn();
	afx_msg void OnBnClickedAddBtn();
	afx_msg void OnBnClickedDelBtn();
	afx_msg void OnItemchangedList1(NMHDR*, LRESULT*);
	CComboBox findCBox;
	CComboBox delCBox;
	CComboBox mfCBox;
	afx_msg void OnBnClickedMfBtn();
	afx_msg void OnCbnSelchangeMfCombo();
	int mfNum;
	//MyCEditEx addStartSN;
	//MyCEditEx mfStartSN;
	CComboBox saleCombo;
	MyCEditEx addSNEdit;
	CEdit mfSNEdit;
	CEdit newSnEdit;
	//CComboBox typeCombo;
	CAutoCombox typeCombo;
	CComboBox statusCombo;
	CComboBox mfStatusCombo;
	CComboBox mfSaleCombo;
};

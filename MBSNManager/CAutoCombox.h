#pragma once
#include <afxwin.h>
class CAutoCombox :
	public CComboBox
{
	DECLARE_DYNAMIC(CAutoCombox)

	public:
		CAutoCombox();
		virtual ~CAutoCombox();
		CString m_model;
		int AddString(LPCTSTR lpszString);
		int DeleteString(UINT nIndex);
		int InsertString(int nIndex, LPCTSTR lpszString);
		void ResetContent();
		void SetFlag(UINT nFlag)
		{
			m_nFlag = nFlag;
		}
		void SetInput(CString model)
		{
			m_model = model;
		}
		CEdit* m_pEdit;//edit control

	private:
		int Dir(UINT attr, LPCTSTR lpszWildCard)
		{
			ASSERT(FALSE);
		}//forbidden

	protected:
		virtual BOOL OnCommand(WPARAM wParam, LPARAM lParam);
		void AutoSelect();
		void AutoMatchAndSel();

		DECLARE_MESSAGE_MAP()


	private:
		UINT m_nFlag;	//some flag
						//bit 0: 0 is show all, 1 is remove not matching, if no maching, show all.
		CStringArray m_strArr;
	
};


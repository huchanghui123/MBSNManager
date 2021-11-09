#include "pch.h"
#include "CAutoCombox.h"


IMPLEMENT_DYNAMIC(CAutoCombox, CComboBox)


CAutoCombox::CAutoCombox()
{
	m_pEdit = NULL;
	m_nFlag = 1;
}

CAutoCombox::~CAutoCombox()
{
	if (m_pEdit)
	{
		if (::IsWindow(m_hWnd))
		{
			m_pEdit->UnsubclassWindow();
		}
		delete m_pEdit;
		m_pEdit = NULL;
	}
}

BEGIN_MESSAGE_MAP(CAutoCombox, CComboBox)
END_MESSAGE_MAP()

void CAutoCombox::AutoSelect()
{
	// Make sure we can 'talk' to the edit control
	if (m_pEdit == NULL)
	{
		m_pEdit = new CEdit();
		m_pEdit->SubclassWindow(GetDlgItem(1001)->GetSafeHwnd());
	}

	// Save the state of the edit control
	CString strText;			//ȡ�������ַ���
	int nStart = 0, nEnd = 0;	//ȡ�ù��λ��
	m_pEdit->GetWindowText(strText);
	m_pEdit->GetSel(nStart, nEnd);

	// Perform actual completion
	int nBestIndex = -1;		//�Ƿ����ҵ�ƥ����ַ�
	int nBestFrom = INT_MAX;	//ƥ�俪ʼ���ַ�

	if (!strText.IsEmpty())
	{
		for (int nIndex = 0; nIndex < GetCount(); ++nIndex)
		{
			CString str;
			GetLBText(nIndex, str);

			int nFrom = str.Find(strText);

			if (nFrom != -1 && nFrom < nBestFrom)//��ƥ�䣬�����Ǹ��õ�ƥ�䣬�ż�¼
			{
				nBestIndex = nIndex;
				nBestFrom = nFrom;
			}
		}
	}

	//Set select index
	if (!GetDroppedState())
	{
		ShowDropDown(TRUE);
		SetCursor(LoadCursor(NULL, MAKEINTRESOURCE(IDC_ARROW)));
		m_pEdit->SetWindowText(strText);
		m_pEdit->SetSel(nStart, nEnd);
	}

	if (GetCurSel() != nBestIndex)
	{
		// Select the matching entry in the list
		SetCurSel(nBestIndex);

		// Restore the edit control
		m_pEdit->SetWindowText(strText);
		m_pEdit->SetSel(nStart, nEnd);
	}

}

void CAutoCombox::AutoMatchAndSel()
{
	// Make sure we can 'talk' to the edit control
	if (m_pEdit == NULL)
	{
		m_pEdit = new CEdit();
		m_pEdit->SubclassWindow(GetDlgItem(1001)->GetSafeHwnd());
	}

	// ����edit�ؼ���״̬
	CString strText;			//ȡ�������ַ���
	int nStart = 0, nEnd = 0;	//ȡ�ù��λ��
	m_pEdit->GetWindowText(strText);
	m_pEdit->GetSel(nStart, nEnd);

	//���CComboBox���������
	CComboBox::ResetContent();

	// ��������б���ѡ������ʵ�
	int nBestIndex = -1;		//�Ƿ����ҵ�ƥ����ַ�
	int nBestFrom = INT_MAX;	//ƥ�俪ʼ���ַ�

	if (!strText.IsEmpty())
	{
		for (int nIndex = 0; nIndex < m_strArr.GetSize(); ++nIndex)
		{
			int nFrom = m_strArr[nIndex].Find(strText);
			char kk = m_strArr[nIndex].GetAt(0);
			char jj = strText.GetAt(0);
			BOOL flag = FALSE;
			if (kk == jj)
			{
				flag = TRUE;
			}
			if (nFrom != -1 && flag == TRUE)//��ƥ��
			{
				int n = CComboBox::AddString(m_strArr[nIndex]);

				if (nFrom < nBestFrom)//���õ�ƥ�䣬���¼
				{
					nBestIndex = n;
					nBestFrom = nFrom;
				}
			}
		}//for
	}

	if (GetCount() == 0)	//û�е���ʾ����
	{
		for (int nIndex = 0; nIndex < m_strArr.GetSize(); ++nIndex)
		{
			CComboBox::AddString(m_strArr[nIndex]);
		}
	}

	//��ʾ�����б�
	if (!GetDroppedState())
	{
		ShowDropDown(TRUE);
		SetCursor(LoadCursor(NULL, MAKEINTRESOURCE(IDC_ARROW)));
	}

	//����ѡ����
	// Select and Restore the edit control
	SetCurSel(nBestIndex);
	m_pEdit->SetWindowText(strText);
	m_pEdit->SetSel(nStart, nEnd);
	
}

int  CAutoCombox::AddString(LPCTSTR lpszString)
{
	m_strArr.Add(lpszString);
	return CComboBox::AddString(lpszString);
}
int  CAutoCombox::DeleteString(UINT nIndex)
{
	m_strArr.RemoveAt(nIndex);
	return CComboBox::DeleteString(nIndex);
}
int  CAutoCombox::InsertString(int nIndex, LPCTSTR lpszString)
{
	m_strArr.InsertAt(nIndex, lpszString);
	return CComboBox::InsertString(nIndex, lpszString);
}
void  CAutoCombox::ResetContent()
{
	m_strArr.RemoveAll();
	CComboBox::ResetContent();
}

//All Message Handle Dispatch
BOOL CAutoCombox::OnCommand(WPARAM wParam, LPARAM lParam)
{
	if (HIWORD(wParam) == EN_CHANGE)
	{
		if (m_nFlag & 0x01)
		{
			AutoMatchAndSel();
		}
		else
		{
			AutoSelect();
		}
		return true;
	}
	else
	{
		return CComboBox::OnCommand(wParam, lParam);
	}
}


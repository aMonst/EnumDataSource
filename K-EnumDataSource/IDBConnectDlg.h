#pragma once
#include "afxwin.h"


// IDBConnectDlg �Ի���

class IDBConnectDlg : public CPropertyPage
{
	DECLARE_DYNAMIC(IDBConnectDlg)

public:
	IDBConnectDlg();
	virtual ~IDBConnectDlg();

// �Ի�������
	enum { IDD = IDD_DLG_CONNECT };

public:
	CComboBox m_ComboDataSource;
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��
	BOOL m_bInit;
	ISourcesRowset* m_pConnSourceRowset;
	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnSetActive();
	void Initialize(const CStringW& csSelected);
	afx_msg void OnCbnDropdownComboDatasource();
	afx_msg void OnBnClickedBtnConnectTest();
};

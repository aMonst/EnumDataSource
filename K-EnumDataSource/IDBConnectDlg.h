#pragma once
#include "afxwin.h"


// IDBConnectDlg 对话框

class IDBConnectDlg : public CPropertyPage
{
	DECLARE_DYNAMIC(IDBConnectDlg)

public:
	IDBConnectDlg();
	virtual ~IDBConnectDlg();

// 对话框数据
	enum { IDD = IDD_DLG_CONNECT };

public:
	CComboBox m_ComboDataSource;
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持
	BOOL m_bInit;
	ISourcesRowset* m_pConnSourceRowset;
	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnSetActive();
	void Initialize(const CStringW& csSelected);
	afx_msg void OnCbnDropdownComboDatasource();
	afx_msg void OnBnClickedBtnConnectTest();
};

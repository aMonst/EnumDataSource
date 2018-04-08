#pragma once
#include "afxwin.h"

// IDBSourceDlg 对话框

class IDBSourceDlg : public CPropertyPage
{
	DECLARE_DYNAMIC(IDBSourceDlg)

public:
	IDBSourceDlg();
	virtual ~IDBSourceDlg();

// 对话框数据
	enum { IDD = IDD_DLG_DATASOUCE };
public:
	CListBox m_DataSourceList;
	CStringW m_csSelected;
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnSetActive();
	void EnumDataSource(ISourcesRowset *pISourceRowset);
	virtual LRESULT OnWizardNext();
};

#pragma once

#include "IDBConnectDlg.h"
#include "IDBSourceDlg.h"

// CDBPropset

class CDBPropset : public CPropertySheet
{
	DECLARE_DYNAMIC(CDBPropset)

public:
	CDBPropset(UINT nIDCaption, CWnd* pParentWnd = NULL, UINT iSelectPage = 0);
	CDBPropset(LPCTSTR pszCaption, CWnd* pParentWnd = NULL, UINT iSelectPage = 0);
	virtual ~CDBPropset();

protected:
	DECLARE_MESSAGE_MAP()
public:
	IDBSourceDlg m_DBSourceDlg;
	IDBConnectDlg m_DBConnectDlg;
};



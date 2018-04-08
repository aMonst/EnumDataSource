// DBPropset.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "K-EnumDataSource.h"
#include "DBPropset.h"


// CDBPropset

IMPLEMENT_DYNAMIC(CDBPropset, CPropertySheet)

CDBPropset::CDBPropset(UINT nIDCaption, CWnd* pParentWnd, UINT iSelectPage)
	:CPropertySheet(nIDCaption, pParentWnd, iSelectPage)
{
	AddPage(&m_DBSourceDlg);
	AddPage(&m_DBConnectDlg);
}

CDBPropset::CDBPropset(LPCTSTR pszCaption, CWnd* pParentWnd, UINT iSelectPage)
	:CPropertySheet(pszCaption, pParentWnd, iSelectPage)
{
	AddPage(&m_DBSourceDlg);
	AddPage(&m_DBConnectDlg);
}

CDBPropset::~CDBPropset()
{
}


BEGIN_MESSAGE_MAP(CDBPropset, CPropertySheet)
END_MESSAGE_MAP()


// CDBPropset ��Ϣ�������

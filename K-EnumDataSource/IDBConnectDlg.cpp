// IDBConnectDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "K-EnumDataSource.h"
#include "IDBConnectDlg.h"
#include "IDBSourceDlg.h"
#include "OLEDBComment.h"

// IDBConnectDlg 对话框

IMPLEMENT_DYNAMIC(IDBConnectDlg, CPropertyPage)

IDBConnectDlg::IDBConnectDlg()
	: CPropertyPage(IDBConnectDlg::IDD)
{
	m_bInit = TRUE;
	m_pConnSourceRowset = NULL;
}

IDBConnectDlg::~IDBConnectDlg()
{
	COM_SAFE_RELEASE(m_pConnSourceRowset);
}

void IDBConnectDlg::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_COMBO_DATASOURCE, m_ComboDataSource);
}


BEGIN_MESSAGE_MAP(IDBConnectDlg, CPropertyPage)
	ON_BN_CLICKED(IDC_BTN_CONNECT_TEST, &IDBConnectDlg::OnBnClickedBtnConnectTest)
END_MESSAGE_MAP()


// IDBConnectDlg 消息处理程序

BOOL IDBConnectDlg::OnSetActive()
{
	// TODO: 在此添加专用代码和/或调用基类
	m_ComboDataSource.ResetContent();
	CPropertySheet *psheet = (CPropertySheet*)GetParent();
	IDBSourceDlg* page = (IDBSourceDlg*)psheet->GetPage(0);
	Initialize(page->m_csSelected);
	psheet->SetWizardButtons(PSWIZB_BACK | PSWIZB_FINISH);
	return CPropertyPage::OnSetActive();
}

void IDBConnectDlg::Initialize(const CStringW& csSelected)
{
	BSTR lpOleName = NULL;
	ULONG uEaten = 0;
	for (vector<DATASOURCE_ENUM_INFO>::iterator it = g_DataSources.begin(); it != g_DataSources.end(); it++)
	{
		if (it->csDisplayName == csSelected)
		{
			lpOleName = it->csSourceName.AllocSysString();
		}
	}

	COM_DECLARE_INTERFACE(ISourcesRowset);
	COM_DECLARE_INTERFACE(IParseDisplayName);
	COM_DECLARE_INTERFACE(IMoniker);

	CoCreateInstance(CLSID_OLEDB_ENUMERATOR, NULL, CLSCTX_INPROC_SERVER, IID_ISourcesRowset, (void**)&pISourcesRowset);
	pISourcesRowset->QueryInterface(IID_IParseDisplayName, (void**)&pIParseDisplayName);
	pIParseDisplayName->ParseDisplayName(NULL, lpOleName, &uEaten, &pIMoniker);
	if (lpOleName != NULL)
	{
		SysFreeString(lpOleName);
	}

	HRESULT hRes = BindMoniker(pIMoniker, 0, IID_ISourcesRowset, (void**)&m_pConnSourceRowset);
	COM_SAFE_RELEASE(pIMoniker);
	COM_SAFE_RELEASE(pIParseDisplayName);

	if (FAILED(hRes))
	{
		COM_SAFE_RELEASE(m_pConnSourceRowset);
		return;
	}
	
	COM_DECLARE_INTERFACE(IRowset)
	hRes = m_pConnSourceRowset->GetSourcesRowset(NULL, IID_IRowset, 0, NULL, (IUnknown**)&pIRowset);
	if (FAILED(hRes))
	{
		return;
	}

	DBBINDING rgBind[1] = {0};
	rgBind[0].bPrecision = 0;
	rgBind[0].bScale = 0;
	rgBind[0].cbMaxLen = 128 * sizeof(WCHAR);
	rgBind[0].dwMemOwner = DBMEMOWNER_CLIENTOWNED;
	rgBind[0].dwPart = DBPART_VALUE;
	rgBind[0].eParamIO = DBPARAMIO_NOTPARAM;
	rgBind[0].iOrdinal = 2; //绑定第二项，用于展示数据源
	rgBind[0].obLength = 0;
	rgBind[0].obStatus = 0;
	rgBind[0].obValue = 0;
	rgBind[0].wType = DBTYPE_WSTR;

	HACCESSOR hAccessor = NULL;
	HROW *rghRows = NULL;
	PVOID pData = NULL;
	PVOID pCurrData = NULL;
	ULONG cRows = 10;
	COM_DECLARE_INTERFACE(IAccessor);
	hRes = pIRowset->QueryInterface(IID_IAccessor, (void**)&pIAccessor);
	if (FAILED(hRes))
	{
		COM_SAFE_RELEASE(pIRowset);
		return;
	}

	hRes = pIAccessor->CreateAccessor(DBACCESSOR_ROWDATA, 1, rgBind, 0, &hAccessor, NULL);
	DBCOUNTITEM cRowsObtained;
	if (FAILED(hRes))
	{
		COM_SAFE_RELEASE(pIRowset);
		COM_SAFE_RELEASE(pIAccessor);
		return;
	}

	pData = MALLOC(rgBind[0].cbMaxLen * cRows);
	while (TRUE)
	{
		hRes = pIRowset->GetNextRows(DB_NULL_HCHAPTER, 0, cRows, &cRowsObtained, &rghRows);
		if (S_OK != hRes && cRowsObtained == 0)
		{
			break;
		}

		for (int i = 0; i < cRowsObtained; i++)
		{
			pCurrData = (BYTE*)pData + rgBind[0].cbMaxLen * i;
			pIRowset->GetData(rghRows[i], hAccessor, pCurrData);

			m_ComboDataSource.AddString((LPOLESTR)pCurrData);
		}

		pIRowset->ReleaseRows(cRowsObtained, rghRows, NULL, NULL, NULL);
	}

	FREE(pData);
	pIAccessor->ReleaseAccessor(hAccessor, NULL);
}

void IDBConnectDlg::OnBnClickedBtnConnectTest()
{
	// TODO: 在此添加控件通知处理程序代码
	CStringW csSelected = _T("");
	ULONG chEaten = 0;
	m_ComboDataSource.GetWindowText(csSelected);
	COM_DECLARE_INTERFACE(IParseDisplayName);
	COM_DECLARE_INTERFACE(IMoniker);
	
	if (m_pConnSourceRowset == NULL)
	{
		MessageBox(_T("连接失败"));
		return;
	}

	HRESULT hRes = m_pConnSourceRowset->QueryInterface(IID_IParseDisplayName, (void**)&pIParseDisplayName);
	if (FAILED(hRes))
	{
		return;
	}

	hRes = pIParseDisplayName->ParseDisplayName(NULL, csSelected.AllocSysString(), &chEaten, &pIMoniker);
	COM_SAFE_RELEASE(pIParseDisplayName);

	if (FAILED(hRes))
	{
		MessageBox(_T("连接失败"));
		return;
	}

	COM_DECLARE_INTERFACE(IDBProperties);
	hRes = BindMoniker(pIMoniker, 0, IID_IDBProperties, (void**)&pIDBProperties);
	COM_SAFE_RELEASE(pIMoniker);

	if (FAILED(hRes))
	{
		MessageBox(_T("连接失败"));
		return;
	}

	//获取用户输入
	CStringW csDB = _T("");
	CStringW csUser = _T("");
	CStringW csPasswd = _T("");
	GetDlgItemText(IDC_EDIT_USERNAME, csUser);
	GetDlgItemText(IDC_EDIT_PASSWORD, csPasswd);
	GetDlgItemText(IDC_EDIT_DATABASE, csDB);

	//设置链接属性
	DBPROP connProp[5] = {0};
	DBPROPSET connPropset[1] = {0};

	connProp[0].colid = DB_NULLID;
	connProp[0].dwOptions = DBPROPOPTIONS_REQUIRED;
	connProp[0].dwPropertyID = DBPROP_INIT_DATASOURCE;
	connProp[0].vValue.vt = VT_BSTR;
	connProp[0].vValue.bstrVal = csSelected.AllocSysString();
	
	connProp[1].colid = DB_NULLID;
	connProp[1].dwOptions = DBPROPOPTIONS_REQUIRED;
	connProp[1].dwPropertyID = DBPROP_INIT_CATALOG;
	connProp[1].vValue.vt = VT_BSTR;
	connProp[1].vValue.bstrVal = csDB.AllocSysString();
	
	connProp[2].colid = DB_NULLID;
	connProp[2].dwOptions = DBPROPOPTIONS_REQUIRED;
	connProp[2].dwPropertyID = DBPROP_AUTH_USERID;
	connProp[2].vValue.vt = VT_BSTR;
	connProp[2].vValue.bstrVal = csUser.AllocSysString();

	connProp[3].colid = DB_NULLID;
	connProp[3].dwOptions = DBPROPOPTIONS_REQUIRED;
	connProp[3].dwPropertyID = DBPROP_AUTH_PASSWORD;
	connProp[3].vValue.vt = VT_BSTR;
	connProp[3].vValue.bstrVal = csPasswd.AllocSysString();

	connPropset[0].cProperties = 4;
	connPropset[0].guidPropertySet = DBPROPSET_DBINIT;
	connPropset[0].rgProperties = connProp;

	hRes = pIDBProperties->SetProperties(1, connPropset);
	if (FAILED(hRes))
	{
		COM_SAFE_RELEASE(pIDBProperties);
		return;
	}

	COM_DECLARE_INTERFACE(IDBInitialize);
	hRes = pIDBProperties ->QueryInterface(IID_IDBInitialize, (void**)&pIDBInitialize);
	COM_SAFE_RELEASE(pIDBProperties);
	if (FAILED(hRes))
	{
		return;
	}

	hRes = pIDBInitialize->Initialize();
	if (FAILED(hRes))
	{
		MessageBox(_T("连接失败"));
	}else
	{
		MessageBox(_T("连接成功"));
	}

	pIDBInitialize->Uninitialize();
	COM_SAFE_RELEASE(pIDBInitialize);
}

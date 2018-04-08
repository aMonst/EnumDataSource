// IDBSourceDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "K-EnumDataSource.h"
#include "IDBSourceDlg.h"
#include "OLEDBComment.h"

// IDBSourceDlg 对话框

IMPLEMENT_DYNAMIC(IDBSourceDlg, CPropertyPage)

IDBSourceDlg::IDBSourceDlg()
	: CPropertyPage(IDBSourceDlg::IDD)
{

}

IDBSourceDlg::~IDBSourceDlg()
{
}

void IDBSourceDlg::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_DataSourceList);
}


BEGIN_MESSAGE_MAP(IDBSourceDlg, CPropertyPage)
END_MESSAGE_MAP()

// IDBSourceDlg 消息处理程序

BOOL IDBSourceDlg::OnSetActive()
{
	// TODO: 在此添加专用代码和/或调用基类
	COM_DECLARE_INTERFACE(ISourcesRowset);
	HRESULT hRes = CoCreateInstance(CLSID_OLEDB_ENUMERATOR, NULL, CLSCTX_INPROC_SERVER, IID_ISourcesRowset, (void**)&pISourcesRowset);
	if (FAILED(hRes))
	{
		ComMessageBox(NULL, _T("OLEDB 错误"), MB_OK, __T("创建接口ISourcesRowset失败,错误码：%08x\n"), hRes);
		goto __CLEAR_UP;
	}
	m_DataSourceList.ResetContent();
	EnumDataSource(pISourcesRowset);

__CLEAR_UP:
	COM_SAFE_RELEASE(pISourcesRowset);
	CPropertySheet *psheet = (CPropertySheet*)GetParent();
	psheet->SetWizardButtons(PSWIZB_NEXT);
	return CPropertyPage::OnSetActive();
}

void IDBSourceDlg::EnumDataSource(ISourcesRowset *pISourceRowset)
{
	COM_DECLARE_INTERFACE(IRowset);
	COM_DECLARE_INTERFACE(IAccessor);
	COM_DECLARE_INTERFACE(IMoniker);
	COM_DECLARE_INTERFACE(IParseDisplayName);
	
	HROW *rgRows = NULL;
	HACCESSOR hAccessor = NULL;
	ULONG cRows = 10;
	DWORD dwOffset = 0;
	PVOID pData = NULL;
	PVOID pCurrentData = NULL;
	DBCOUNTITEM cRowsObtained = 0;
	LPOLESTR lpParamName = OLESTR("");
	//利用顶层枚举对象来枚举系统中存在的数据源
	HRESULT hRes = pISourceRowset->GetSourcesRowset(NULL, IID_IRowset, 0, NULL, (IUnknown**)&pIRowset);
	ISourcesRowset *pSubSourceRowset = NULL;
	ULONG ulEaten = 0;
	if (FAILED(hRes))
	{
		ComMessageBox(NULL, _T("OLEDB 错误"), MB_OK, __T("创建接口ISourcesRowset失败,错误码：%08x\n"), hRes);
		goto __CLEAR_UP;
	}

	DBBINDING rgBinding[3] = {0}; //这里只关心我们需要的列，不需要获取所有的列
	for (int i = 0; i < 3; i++)
	{
		rgBinding[i].bPrecision = 0;
		rgBinding[i].bScale = 0;
		rgBinding[i].cbMaxLen = 128 * sizeof(WCHAR);
		rgBinding[i].dwMemOwner = DBMEMOWNER_CLIENTOWNED;
		rgBinding[i].dwPart = DBPART_VALUE;
		rgBinding[i].eParamIO = DBPARAMIO_NOTPARAM;
		rgBinding[i].wType = DBTYPE_WSTR;
		rgBinding[i].obLength = 0;
		rgBinding[i].obStatus = 0;
		rgBinding[i].obValue = dwOffset;
		dwOffset += rgBinding[0].cbMaxLen;
	}
	rgBinding[0].iOrdinal = 3; //第三项是数据源或者枚举器的描述信息,用于显示
	rgBinding[1].wType = DBTYPE_UI2;
	rgBinding[1].iOrdinal = 4; //第四列是枚举出来的类型信息,用于判断是否需要递归
	rgBinding[2].iOrdinal = 1; //第一列是枚举出来的类型信息，用于获取子枚举器

	hRes = pIRowset->QueryInterface(IID_IAccessor, (void**)&pIAccessor);
	if (FAILED(hRes))
	{
		ComMessageBox(NULL, _T("OLEDB 错误"), MB_OK, __T("查询接口pIAccessor失败,错误码：%08x\n"), hRes);
		goto __CLEAR_UP;
	}
	hRes = pIAccessor->CreateAccessor(DBACCESSOR_ROWDATA, 3, rgBinding, 0, &hAccessor, NULL);
	if (FAILED(hRes))
	{
		ComMessageBox(NULL, _T("OLEDB 错误"), MB_OK, __T("创建访问器失败,错误码：%08x\n"), hRes);
		goto __CLEAR_UP;
	}
	pData = MALLOC(dwOffset * cRows);

	while (TRUE)
	{
		hRes = pIRowset->GetNextRows(DB_NULL_HCHAPTER, 0, cRows, &cRowsObtained, &rgRows);
		if(S_OK != hRes && cRowsObtained == 0)
		{
			break;
		}

		ZeroMemory(pData, dwOffset * cRows);
		for (int i = 0; i < cRowsObtained; i++)
		{
			pCurrentData = (BYTE*)pData + dwOffset * i;
			pIRowset->GetData(rgRows[i], hAccessor, pCurrentData);
			DATASOURCE_ENUM_INFO dbei = {0}; //将枚举到的相关信息存储到对应的结构中
			dbei.csSourceName = (LPCTSTR)((BYTE*)pCurrentData + rgBinding[2].obValue);
			dbei.csDisplayName = (LPCTSTR)((BYTE*)pCurrentData + rgBinding[0].obValue);
			dbei.dbTypeEnum = *(DBTYPEENUM*)((BYTE*)pCurrentData + rgBinding[1].obValue);
			m_DataSourceList.AddString(dbei.csDisplayName); //显示数据源信息
			g_DataSources.push_back(dbei);
		}

		pIRowset->ReleaseRows(cRowsObtained, rgRows, NULL, NULL, NULL);
	}

	pIAccessor->ReleaseAccessor(hAccessor, NULL);

__CLEAR_UP:
	FREE(pData);
	CoTaskMemFree(rgRows);
	COM_SAFE_RELEASE(pIRowset);
	COM_SAFE_RELEASE(pIAccessor);
}

LRESULT IDBSourceDlg::OnWizardNext()
{
	// TODO: 在此添加专用代码和/或调用基类
	int nCurSel = m_DataSourceList.GetCurSel();
	if (nCurSel == -1)
	{
		MessageBox(_T("请选择数据源提供程序")_T("警告"));
		return -1;
	}

	m_DataSourceList.GetText(nCurSel, m_csSelected);
	return CPropertyPage::OnWizardNext();
}

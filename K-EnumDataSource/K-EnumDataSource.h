
// K-EnumDataSource.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CKEnumDataSourceApp:
// �йش����ʵ�֣������ K-EnumDataSource.cpp
//

class CKEnumDataSourceApp : public CWinAppEx
{
public:
	CKEnumDataSourceApp();

// ��д
	public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CKEnumDataSourceApp theApp;
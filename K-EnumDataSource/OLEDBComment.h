#pragma once;

#include <oledb.h>
#include <oledberr.h>
#include <tchar.h>
#include <strsafe.h>
#include <msdasc.h>
#include <msdaguid.h>
#include <vector>
using namespace std;

#define COM_DECLARE_INTERFACE(I) I* p##I = NULL;

int ComMessageBox(HWND hWnd, LPTSTR lpCaption, UINT uType, LPTSTR lpszFormat, ...);

#define COM_SAFE_RELEASE(I)\
	if(NULL != (I))\
	{\
		(I)->Release();\
		(I) = NULL;\
	}

#define MALLOC(s) HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, s)
#define FREE(p) HeapFree(GetProcessHeap(), 0, p)

typedef struct _tag_DATASOURCE_ENUM_INFO
{
	CStringW csSourceName;
	CStringW csDisplayName;
	DBTYPEENUM dbTypeEnum;
}DATASOURCE_ENUM_INFO, *LPDATASOURCE_ENUM_INFO;

extern vector<DATASOURCE_ENUM_INFO> g_DataSources;
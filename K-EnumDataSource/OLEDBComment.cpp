#include "stdafx.h"
#include "OLEDBComment.h"

int ComMessageBox(HWND hWnd, LPTSTR lpCaption, UINT uType, LPTSTR lpszFormat, ...)
{
	TCHAR szBuf[1024] = _T("");
	va_list pargList;
	va_start(pargList, lpszFormat);
	StringCchPrintf(szBuf, 1024, lpszFormat, pargList);
	va_end(pargList);

	return MessageBox(hWnd, szBuf, lpCaption, uType);
}
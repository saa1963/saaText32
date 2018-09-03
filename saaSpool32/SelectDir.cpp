#include "StdAfx.h"
#include <windows.h>
#include <atlbase.h>
#include <shlobj.h>
#pragma comment(lib,"shell32")

static int CALLBACK
BrowseCallbackProc(HWND hWnd, UINT uMsg, LPARAM lParam, LPARAM lpData)
{
	TCHAR szPath[_MAX_PATH];
	switch (uMsg) {
	case BFFM_INITIALIZED:
		if (lpData)
			SendMessage(hWnd, BFFM_SETSELECTION, TRUE, lpData);
		break;
	case BFFM_SELCHANGED:
		SHGetPathFromIDList(LPITEMIDLIST(lParam), szPath);
		SendMessage(hWnd, BFFM_SETSTATUSTEXT, NULL, LPARAM(szPath));
		break;
	}
	return 0;
}

BOOL GetFolder(LPCTSTR szTitle, LPTSTR szPath, LPCTSTR szRoot, HWND hWndOwner)
{
	if (szPath == NULL)
		return false;

	bool result = false;

	LPMALLOC pMalloc;
	if (::SHGetMalloc(&pMalloc) == NOERROR) {
		BROWSEINFO bi;
		::ZeroMemory(&bi, sizeof bi);
		bi.ulFlags = BIF_RETURNONLYFSDIRS;

		// дескриптор окна-владельца диалога
		bi.hwndOwner = hWndOwner;

		// добавление заголовка к диалогу
		bi.lpszTitle = szTitle;

		// отображение текущего каталога
		bi.lpfn = BrowseCallbackProc;
		bi.ulFlags |= BIF_STATUSTEXT;

		// установка каталога по умолчанию
		bi.lParam = LPARAM(szPath);

		// установка корневого каталога
		if (szRoot != NULL) {
			IShellFolder *pDF;
			if (SHGetDesktopFolder(&pDF) == NOERROR) {
				LPITEMIDLIST pIdl = NULL;
				ULONG        chEaten;
				ULONG        dwAttributes;

				USES_CONVERSION;
				LPOLESTR oleStr = T2OLE(szRoot);

				pDF->ParseDisplayName(NULL, NULL, oleStr, &chEaten, &pIdl, &dwAttributes);
				pDF->Release();

				bi.pidlRoot = pIdl;
			}
		}

		LPITEMIDLIST pidl = ::SHBrowseForFolder(&bi);
		if (pidl != NULL) {
			if (::SHGetPathFromIDList(pidl, szPath))
				result = true;
			pMalloc->Free(pidl);
		}
		if (bi.pidlRoot != NULL)
			pMalloc->Free((void*)bi.pidlRoot);
		pMalloc->Release();
	}
	return result;
}


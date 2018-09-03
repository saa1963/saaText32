// CSpooler.cpp : Implementation of CCSpooler

#include "stdafx.h"
#include "CSpooler.h"
#include ".\cspooler.h"
#include "dtxview.h"
#include <vector>
#include <set>
#include <string>
#include "SaaPrint.h"

using namespace std;
#pragma warning(disable : 4996)
#pragma warning(disable : 4018)
// CCSpooler

struct RETURN_STRUC {
	char *pName;
};

BOOL CALLBACK SelectPrinterProc(HWND hwndDlg, UINT message, WPARAM wParam, LPARAM lParam);
int FillListBox(HWND hwndDlg);

bool GetPrinterName(string& s) {
	RETURN_STRUC *rt = NULL;
	LPCTSTR lpTemplate = MAKEINTRESOURCE(IDD_SELECTPRINTER);
	DWORD n = (int)DialogBoxParam(GetModuleHandle("saaSpool"), lpTemplate, NULL,
		(DLGPROC)SelectPrinterProc, (LPARAM)NULL);
	rt = (RETURN_STRUC*)n;
	if (rt != 0) {
		s.assign((*rt).pName);
		delete (*rt).pName;
		delete rt;
		return true;
	}
	else
		return false;
}

BOOL CALLBACK SelectPrinterProc(HWND hwndDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	char *s = NULL;
	int selitem;
	HWND hwndOwner;
	RECT rc, rcDlg, rcOwner;
	RETURN_STRUC *rt;

	switch (message)
	{

		// Place message cases here. 
	case WM_INITDIALOG:
		// Get the owner window and dialog box rectangles. 

		if ((hwndOwner = GetParent(hwndDlg)) == NULL)
		{
			hwndOwner = GetDesktopWindow();
		}

		GetWindowRect(hwndOwner, &rcOwner);
		GetWindowRect(hwndDlg, &rcDlg);
		CopyRect(&rc, &rcOwner);

		// Offset the owner and dialog box rectangles so that 
		// right and bottom values represent the width and 
		// height, and then offset the owner again to discard 
		// space taken up by the dialog box. 

		OffsetRect(&rcDlg, -rcDlg.left, -rcDlg.top);
		OffsetRect(&rc, -rc.left, -rc.top);
		OffsetRect(&rc, -rcDlg.right, -rcDlg.bottom);

		// The new position is the sum of half the remaining 
		// space and the owner's original position. 

		SetWindowPos(hwndDlg,
			HWND_TOP,
			rcOwner.left + (rc.right / 2),
			rcOwner.top + (rc.bottom / 2),
			0, 0,          // ignores size arguments 
			SWP_NOSIZE);

		FillListBox(hwndDlg);
		return TRUE;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDC_LIST1:
			EnableWindow(GetDlgItem(hwndDlg, IDOK), TRUE);
			return TRUE;
		case IDOK:
			selitem = SendDlgItemMessage(hwndDlg, IDC_LIST1, LB_GETCURSEL, 0, 0);
			rt = new RETURN_STRUC;
			(*rt).pName = new char[SendDlgItemMessage(hwndDlg, IDC_LIST1, LB_GETTEXTLEN, selitem, 0) + 1];
			SendDlgItemMessage(hwndDlg, IDC_LIST1, LB_GETTEXT, selitem, (LPARAM)(*rt).pName);
			EndDialog(hwndDlg, (INT_PTR)(rt));
			return TRUE;
		case IDCANCEL:
			EndDialog(hwndDlg, 0);
			return TRUE;
		default:
			return FALSE;
		}
	case WM_CLOSE:
		EndDialog(hwndDlg, 0);
		return TRUE;
	default:
		return FALSE;
	}
}

int FillListBox(HWND hwndDlg) {
	DWORD cbNeeded, cReturned;
	PRINTER_INFO_2 *info2 = NULL;
	EnumPrinters(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, NULL, 2, NULL, 0, &cbNeeded, &cReturned);
	info2 = (PRINTER_INFO_2*)new BYTE[cbNeeded];
	EnumPrinters(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, NULL, 2, (LPBYTE)info2, cbNeeded, &cbNeeded, &cReturned);

	for (DWORD i = 0; i < cReturned; i++) {
		SendDlgItemMessage(hwndDlg, IDC_LIST1, LB_ADDSTRING, 0, (LPARAM)info2[i].pPrinterName);
	}
	delete[] info2;

	return 1;
}

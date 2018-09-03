// saaSpool32.cpp: определ€ет точку входа дл€ приложени€.
//

#include "stdafx.h"
#include "saaSpool32.h"
#include "DtxView.h"

// ќтправить объ€влени€ функций, включенных в этот модуль кода:

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

	char *lpCmdLineMultiByte;
	CDtxView dv(hInstance);
	if (wcscmp(lpCmdLine, L"") == 0)
		dv.Show();
	else
	{
		int cbMultiByte = WideCharToMultiByte(CP_ACP, 0, lpCmdLine, -1, NULL, 0, NULL, NULL);
		lpCmdLineMultiByte = new char[cbMultiByte];
		WideCharToMultiByte(CP_ACP, 0, lpCmdLine, -1, lpCmdLineMultiByte, cbMultiByte, NULL, NULL);
		dv.Show(lpCmdLineMultiByte);
		delete lpCmdLineMultiByte;
	}
    return 0;
}

#include "StdAfx.h"
#include "windowsx.h"
#include "OaIdl.h"
#include "Commdlg.h"
#include ".\dtxview.h"
#include "resource.h"
#include "commctrl.h"
#include "selectdir.h"
#include <vector>
#include <string>
#include <map>
#include <regex>
#include "saacutils.h"
#include "CSpooler.h"
#include "SaaPrint.h"

using namespace std;

#define CX_ICON  32 
#define CY_ICON  32 
#define NUM_ICONS 6
#define MSG_FINDFIRST 10000
#define MSG_FINDNEXT 10001
#define MAXEDITBUF 10240000
#pragma warning(disable : 4996)
#pragma warning(disable : 4018)


BOOL CALLBACK DtxDialogProc(HWND hwndDlg, UINT message, WPARAM wParam, LPARAM lParam);
void ResizeDialog(CDtxView *dtxview, HWND hwndDlg, int lbWidth);
void CreateTB(CDtxView *dtxview, HWND hwndDlg);
BOOL CreateToolTip(LPTOOLTIPTEXT lpttt);
void CenterDialog(HWND hwndDlg);
void SavePosition(HWND hwndDlg);
void WordText(string& s);
BOOL RestPosition(HWND hwndDlg);
void SaveFile(HWND hwndDlg, const string& fileName);
void OpenDtx(HWND hwndDlg);
void ReFill(HWND hwndDlg);
void OpenDtxFile(HWND hwndDlg, char *fname);
void GetListFiles(char *dir, vector<string> &files);
void ReleaseItemData(HWND hwndOpenDoc);
BOOL CALLBACK FindProc(HWND hwndFind, UINT message, WPARAM wParam, LPARAM lParam);
void FindFirst(HWND hwndDlg, string *find);
void FindNext(HWND hwndDlg, string *find);
void ErrHandler(HRESULT hr, EXCEPINFO excep);
WNDPROC pOldProc;

typedef pair<string, string> PairString;

CDtxView::CDtxView(HINSTANCE hInst)
	: option(NULL)
	, CurSel(0)
	, IsChanged(FALSE)
	, findpos(0)
	, page1(0)
	, page2(0)
	, hInstance(hInst)
{

	//HINSTANCE hInst = GetModuleHandle("saaSpool");
	//HINSTANCE hInst = _AtlBaseModule.GetModuleInstance();
	HICON hIcon;
	INITCOMMONCONTROLSEX icex;
	ZeroMemory(&icex, sizeof(INITCOMMONCONTROLSEX));

	hEditFont = this->GetFontForEditBox();
	// Ensure that the common control DLL is loaded. 
	icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
	icex.dwICC = ICC_BAR_CLASSES;
	InitCommonControlsEx(&icex);

	// Create a masked image list large enough to hold the icons. 
	hImageList = ImageList_Create(CX_ICON, CY_ICON, ILC_COLOR16 | ILC_MASK, NUM_ICONS, 0);

	// Load the icon resources, and add the icons to the image list. 
	hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_OPENFILE));
	ImageList_AddIcon(hImageList, hIcon);
	hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_PRINTALL));
	ImageList_AddIcon(hImageList, hIcon);
	hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_EDIT));
	ImageList_AddIcon(hImageList, hIcon);
	hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_FIND));
	ImageList_AddIcon(hImageList, hIcon);
	hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_MSWORD));
	ImageList_AddIcon(hImageList, hIcon);
	hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_QUIT));
	ImageList_AddIcon(hImageList, hIcon);
}

CDtxView::~CDtxView(void)
{
	ImageList_Destroy(hImageList);
	DeleteObject(hEditFont);
}

void CDtxView::Show(char* fname)
{
	//dtx = new CDtx();
	dtx = NULL;
	try
	{
		this->text = readFile(fname);
	//if (dtx->Open(fname)) {
		
	//	dtx->Close();
	}
	catch(...)
	{
		::MessageBox(NULL, "Ошибка открытия файла", "", MB_OK);
		return;
	}
	_Show();
	//delete dtx;
}

std::string CDtxView::readFile(const std::string& fileName) {
	std::ifstream t(fileName, ifstream::binary);
	std::string str;
	std::string empty("");
	std::regex rx("`(.*)`");

	t.seekg(0, std::ios::end);
	str.reserve(t.tellg());
	t.seekg(0, std::ios::beg);

	str.assign((std::istreambuf_iterator<char>(t)),
		std::istreambuf_iterator<char>());
	return std::regex_replace(str, rx, empty);
}

bool CDtxView::writeFile(const std::string& fileName) {
	std::ofstream out;
	bool rt = true;
	try
	{
		out.open(fileName, ofstream::binary); // окрываем файл для записи
		if (out.is_open())
		{
			out << this->text;
		}
	}
	catch (...)
	{
		rt = false;
		::MessageBox(NULL, "Ошибка записи файла.", "", MB_OK);
	}
	if (out.is_open())
		out.close();     // закрываем файл
	return rt;
}

void CDtxView::Show(void)
{
	dtx = NULL;
	_Show();
}

void CDtxView::_Show()
{
	LPCTSTR lpTemplate = MAKEINTRESOURCE(IDD_DTX);
	InitCommonControls();
	DWORD n = (int)DialogBoxParam(hInstance, lpTemplate, NULL, (DLGPROC)DtxDialogProc, (LPARAM)this);
}

BOOL CALLBACK DtxDialogProc(HWND hwndDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	string temp;
	CDtxView *dtxview;
	int selcount;
	string pName;
	int *selitems;
	CSaaPrint *print;

	switch (message)
	{
		// Place message cases here. 
	case WM_INITDIALOG:
		// сохраняем указатель на объект CDtxView
		SetWindowLong(hwndDlg, DWL_USER, (LONG)lParam);

		// Устанавливаем шрифт для EDIT
		SendMessage(GetDlgItem(hwndDlg, IDC_EDIT1), WM_SETFONT, (WPARAM)((CDtxView*)lParam)->hEditFont, (LPARAM)0);

		// устанавливаем максимальный размер буфера
		SendMessage(GetDlgItem(hwndDlg, IDC_EDIT1), EM_SETLIMITTEXT, (WPARAM)MAXEDITBUF, (LPARAM)0);

		// Создаем TOOLBAR CDtxView::hwndTB
		CreateTB((CDtxView*)lParam, hwndDlg);

		// центрируем диалог
		if (!RestPosition(hwndDlg))
			CenterDialog(hwndDlg);

		ReFill(hwndDlg);

		return TRUE;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDC_EDIT1:
			switch (HIWORD(wParam)) {
			case EN_CHANGE:
				dtxview = (CDtxView*)GetWindowLong(hwndDlg, DWL_USER);
				dtxview->IsChanged = TRUE;
				return TRUE;
			default:
				return FALSE;
			}
			return TRUE;
		case IDI_OPENFILE:
			OpenDtx(hwndDlg);
			return TRUE;
		case IDI_PRINTALL:
			dtxview = (CDtxView*)GetWindowLong(hwndDlg, DWL_USER);
			if (GetPrinterName(pName)) {
				print = new CSaaPrint(pName);
				//print->PrintText(dtxview->text, false);
				print->PrintText(GetEditText(GetDlgItem(hwndDlg, IDC_EDIT1)), false);
				delete print;
			}
			return TRUE;
		case IDI_EDIT:
			dtxview = (CDtxView*)GetWindowLong(hwndDlg, DWL_USER);
			temp = GetEditText(GetDlgItem(hwndDlg, IDC_EDIT1));
			SaveFile(hwndDlg, temp);
			dtxview->IsChanged = FALSE;
			return TRUE;
		case IDI_FIND:
			dtxview = (CDtxView*)GetWindowLong(hwndDlg, DWL_USER);
			dtxview->hwndFind = CreateDialog(dtxview->hInstance, MAKEINTRESOURCE(IDD_FIND), hwndDlg, FindProc);
			ShowWindow(dtxview->hwndFind, SW_SHOWNORMAL);
			return TRUE;
		case IDI_MSWORD:
			dtxview = (CDtxView*)GetWindowLong(hwndDlg, DWL_USER);
			WordText(GetEditText(GetDlgItem(hwndDlg, IDC_EDIT1)));
			return TRUE;
		case IDI_QUIT:
			EndDialog(hwndDlg, 0);
			return TRUE;
		default:
			return FALSE;
		}
	case WM_SIZE:
		dtxview = (CDtxView*)GetWindowLong(hwndDlg, DWL_USER);
		ResizeDialog(dtxview, hwndDlg, 250);
		return FALSE;
	case WM_NOTIFY:
		switch (((LPNMHDR)lParam)->code)
		{
		case TTN_NEEDTEXTA:
			return CreateToolTip((LPTOOLTIPTEXT)lParam);
		default:
			return FALSE;
		}
		return FALSE;
	case MSG_FINDFIRST:
		FindFirst(hwndDlg, (string*)lParam);
		return TRUE;
	case MSG_FINDNEXT:
		FindNext(hwndDlg, (string*)lParam);
		return TRUE;
	case WM_CLOSE:
		EndDialog(hwndDlg, 0);
		return TRUE;
	case WM_DESTROY:
		// запомним расположение окна
		SavePosition(hwndDlg);
		return FALSE;
	default:
		return FALSE;
	}
}

HFONT CDtxView::GetFontForEditBox(void)
{
	LOGFONT lf;
	ZeroMemory(&lf, sizeof(LOGFONT));
	lf.lfHeight = 14;
	lf.lfWidth = 0;
	lf.lfEscapement = 0;
	lf.lfOrientation = 0;
	lf.lfWeight = 0;
	lf.lfItalic = FALSE;
	lf.lfUnderline = FALSE;
	lf.lfStrikeOut = FALSE;
	lf.lfCharSet = DEFAULT_CHARSET;
	lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
	lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
	lf.lfQuality = DEFAULT_QUALITY;
	lf.lfPitchAndFamily = 0;
	strcpy(lf.lfFaceName, "Courier New");
	return CreateFontIndirect(&lf);
}

void ResizeDialog(CDtxView *dtxview, HWND hwndDlg, int lbWidth) {
	RECT rect;
	int heightToolbar;
	int dlgWidth, dlgHeight, dlgLeft, dlgRight, dlgTop, dlgBottom;

	SendMessage(dtxview->hwndTB, TB_AUTOSIZE, 0, 0);

	// Определяем высоту TOOLBAR
	GetWindowRect(dtxview->hwndTB, &rect);
	heightToolbar = rect.bottom - rect.top + 1;

	// размеры клиентской части диалога
	GetClientRect(hwndDlg, &rect);
	dlgLeft = rect.left; dlgTop = rect.top;
	dlgRight = rect.right; dlgBottom = rect.bottom;
	dlgHeight = rect.bottom - rect.top + 1;
	dlgWidth = rect.right - rect.left + 1;

	GetWindowRect(GetDlgItem(hwndDlg, IDC_EDIT1), &rect);
	//MoveWindow(GetDlgItem(hwndDlg, IDC_EDIT1), lbWidth + stWidth, heightToolbar, dlgWidth - lbWidth - stWidth, dlgHeight - heightToolbar, TRUE);
	MoveWindow(GetDlgItem(hwndDlg, IDC_EDIT1), dlgLeft, heightToolbar, dlgWidth, dlgHeight - heightToolbar, TRUE);
}

void CreateTB(CDtxView *dtxview, HWND hwndDlg) {
	HWND hwndTB;
	TBBUTTON aBtn[NUM_ICONS];
	ZeroMemory(aBtn, NUM_ICONS * sizeof(TBBUTTON));
	// создаем TOOLBAR
	hwndTB = dtxview->hwndTB = CreateWindowEx(0, TOOLBARCLASSNAME, (LPSTR)NULL,
		WS_CHILD | CCS_ADJUSTABLE | WS_VISIBLE | TBSTYLE_TOOLTIPS, 0, 0, 0, 0, hwndDlg,
		0, 0, NULL);
	// Send the TB_BUTTONSTRUCTSIZE message, which is required for 
	// backward compatibility. 
	SendMessage(hwndTB, TB_BUTTONSTRUCTSIZE, (WPARAM) sizeof(TBBUTTON), 0);
	// для возможности добавлять несколько ImageList
	//SendMessage(hwndTB, CCM_SETVERSION, (WPARAM) 5, 0); 
	// привязываем ImageList к Toolbar
	SendMessage(hwndTB, TB_SETIMAGELIST, (WPARAM)0, (LPARAM)dtxview->hImageList);

	aBtn[0].iBitmap = MAKELONG(0, 0);
	aBtn[0].idCommand = IDI_OPENFILE;
	aBtn[0].fsState = TBSTATE_ENABLED;
	aBtn[0].fsStyle = TBSTYLE_BUTTON;

	aBtn[1].iBitmap = MAKELONG(1, 0);
	aBtn[1].idCommand = IDI_PRINTALL;
	aBtn[1].fsState = TBSTATE_ENABLED;
	aBtn[1].fsStyle = TBSTYLE_BUTTON;

	aBtn[2].iBitmap = MAKELONG(2, 0);
	aBtn[2].idCommand = IDI_EDIT;
	aBtn[2].fsState = TBSTATE_ENABLED;
	aBtn[2].fsStyle = TBSTYLE_BUTTON;

	aBtn[3].iBitmap = MAKELONG(3, 0);
	aBtn[3].idCommand = IDI_FIND;
	aBtn[3].fsState = TBSTATE_ENABLED;
	aBtn[3].fsStyle = TBSTYLE_BUTTON;

	aBtn[4].iBitmap = MAKELONG(4, 0);
	aBtn[4].idCommand = IDI_MSWORD;
	aBtn[4].fsState = TBSTATE_ENABLED;
	aBtn[4].fsStyle = TBSTYLE_BUTTON;

	aBtn[5].iBitmap = MAKELONG(5, 0);
	aBtn[5].idCommand = IDI_QUIT;
	aBtn[5].fsState = TBSTATE_ENABLED;
	aBtn[5].fsStyle = TBSTYLE_BUTTON;

	// Add the five buttons to the toolbar control
	SendMessage(hwndTB, TB_ADDBUTTONS, NUM_ICONS, (LPARAM)(&aBtn));
	SendMessage(hwndTB, TB_AUTOSIZE, 0, 0);
}

BOOL CreateToolTip(LPTOOLTIPTEXT lpttt) {
	lpttt->hinst = NULL;
	switch (lpttt->hdr.idFrom) {
	case IDI_OPENFILE:
		lpttt->lpszText = "Открыть файл";
		return TRUE;
	case IDI_PRINTALL:
		lpttt->lpszText = "Печатать";
		return TRUE;
	case IDI_EDIT:
		lpttt->lpszText = "Сохранить изменения";
		return TRUE;
	case IDI_FIND:
		lpttt->lpszText = "Поиск";
		return TRUE;
	case IDI_MSWORD:
		lpttt->lpszText = "Передать в MS Word";
		return TRUE;
	case IDI_QUIT:
		lpttt->lpszText = "Выход";
		return TRUE;
	default:
		return FALSE;
	}
}

void CenterDialog(HWND hwndDlg) {
	HWND hwndOwner;
	RECT rc, rcDlg, rcOwner;
	// Get the owner window and dialog box rectangles. 

	if ((hwndOwner = GetParent(hwndDlg)) == NULL)
	{
		hwndOwner = GetDesktopWindow();
	}

	GetWindowRect(hwndOwner, &rcOwner);
	GetWindowRect(hwndDlg, &rcDlg);
	MoveWindow(hwndDlg, rcDlg.left, rcDlg.top, rcDlg.right - rcDlg.left + 1, rcDlg.bottom - rcDlg.top, TRUE);
	CopyRect(&rc, &rcOwner);

	//// Offset the owner and dialog box rectangles so that 
	//// right and bottom values represent the width and 
	//// height, and then offset the owner again to discard 
	//// space taken up by the dialog box. 

	OffsetRect(&rcDlg, -rcDlg.left, -rcDlg.top);
	OffsetRect(&rc, -rc.left, -rc.top);
	OffsetRect(&rc, -rcDlg.right, -rcDlg.bottom);

	//// The new position is the sum of half the remaining 
	//// space and the owner's original position. 

	SetWindowPos(hwndDlg,
		HWND_TOP,
		rcOwner.left + (rc.right / 2),
		rcOwner.top + (rc.bottom / 2),
		0, 0,          // ignores size arguments 
		SWP_NOSIZE);
}

void SavePosition(HWND hwndDlg) {
	RECT rect;
	RECT rectlb;
	ZeroMemory(&rect, sizeof(RECT));
	ZeroMemory(&rectlb, sizeof(RECT));
	BOOL bMax = FALSE;
	HKEY hKey;
	DWORD data = 0;

	GetWindowRect(hwndDlg, &rect);
	bMax = IsZoomed(hwndDlg);
	GetWindowRect(GetDlgItem(hwndDlg, IDC_LB), &rectlb);

	if (RegCreateKeyEx(HKEY_CURRENT_USER, "Software\\saaText", 0, 0, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, &hKey, 0) == ERROR_SUCCESS) {
		data = (DWORD)rect.left;
		RegSetValueEx(hKey, "left", 0, REG_DWORD, (const BYTE*)&data, 4);
		data = (DWORD)rect.top;
		RegSetValueEx(hKey, "top", 0, REG_DWORD, (const BYTE*)&data, 4);
		data = (DWORD)rect.right;
		RegSetValueEx(hKey, "right", 0, REG_DWORD, (const BYTE*)&data, 4);
		data = (DWORD)rect.bottom;
		RegSetValueEx(hKey, "bottom", 0, REG_DWORD, (const BYTE*)&data, 4);
		data = (DWORD)bMax;
		RegSetValueEx(hKey, "max", 0, REG_DWORD, (const BYTE*)&data, 4);
		data = (DWORD)(rectlb.right - rectlb.left + 1);
		RegSetValueEx(hKey, "widthlb", 0, REG_DWORD, (const BYTE*)&data, 4);
		RegCloseKey(hKey);
	}
	else
		MessageBox(NULL, "Ошибка сохранения данных", "", MB_OK);
}

BOOL RestPosition(HWND hwndDlg) {
	HKEY hKey;
	DWORD data = 0;
	DWORD size = 4;
	DWORD type = REG_DWORD;
	LONG left = 0, top = 0, right = 0, bottom = 0, widthlb = 0, bMax = 0;
	if (RegOpenKeyEx(HKEY_CURRENT_USER, "Software\\saaText", 0, KEY_ALL_ACCESS, &hKey) == ERROR_SUCCESS) {
		RegQueryValueEx(hKey, "left", 0, &type, (LPBYTE)&data, &size);
		left = (LONG)data;
		RegQueryValueEx(hKey, "top", 0, &type, (LPBYTE)&data, &size);
		top = (LONG)data;
		RegQueryValueEx(hKey, "right", 0, &type, (LPBYTE)&data, &size);
		right = (LONG)data;
		RegQueryValueEx(hKey, "bottom", 0, &type, (LPBYTE)&data, &size);
		bottom = (LONG)data;
		RegQueryValueEx(hKey, "widthlb", 0, &type, (LPBYTE)&data, &size);
		widthlb = (LONG)data;
		RegQueryValueEx(hKey, "max", 0, &type, (LPBYTE)&data, &size);
		bMax = (LONG)data;
		if (bMax == 1) {
			ShowWindow(hwndDlg, SW_MAXIMIZE);
			ResizeDialog((CDtxView*)GetWindowLong(hwndDlg, DWL_USER), hwndDlg, widthlb);
		}
		else {
			MoveWindow(hwndDlg, left, top, right - left + 1, bottom - top + 1, TRUE);
			ResizeDialog((CDtxView*)GetWindowLong(hwndDlg, DWL_USER), hwndDlg, widthlb);
		}
		RegCloseKey(hKey);
		return TRUE;
	}
	else
		return FALSE;
}

void OpenDtx(HWND hwndDlg) {
	OPENFILENAME of;
	TCHAR strFile[500];
	strFile[0] = '\0';
	char *filter = "Все файлы\0*.*\0\0";
	ZeroMemory(&of, sizeof(OPENFILENAME));
	of.lStructSize = sizeof(OPENFILENAME);
	of.hwndOwner = hwndDlg;
	of.lpstrFilter = filter;
	of.Flags = OFN_HIDEREADONLY;
	of.lpstrFile = strFile;
	of.nMaxFile = 500;
	//of.lpstrCustomFilter = filter;
	if (GetOpenFileName(&of)) {
		OpenDtxFile(hwndDlg, of.lpstrFile);
	}
}

void SaveFile(HWND hwndDlg, const string& fileName)
{
	CDtxView *dtxview = (CDtxView*)GetWindowLong(hwndDlg, DWL_USER);
	OPENFILENAME of;
	TCHAR strFile[500];
	strFile[0] = '\0';
	char *filter = "Все файлы\0*.*\0\0";
	ZeroMemory(&of, sizeof(OPENFILENAME));
	of.lStructSize = sizeof(OPENFILENAME);
	of.hwndOwner = hwndDlg;
	of.lpstrFilter = filter;
	of.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
	of.lpstrFile = strFile;
	of.nMaxFile = MAX_PATH;
	//of.lpstrCustomFilter = filter;
	if (GetSaveFileName(&of)) {
		dtxview->writeFile(of.lpstrFile);
	}
}

void OpenDtxFile(HWND hwndDlg, char *fname) {
	CDtxView *dtxview = (CDtxView*)GetWindowLong(hwndDlg, DWL_USER);
	dtxview->text = dtxview->readFile(fname);
	ReFill(hwndDlg);
}

void ReFill(HWND hwndDlg) {
	CDtxView *dtxview = (CDtxView*)GetWindowLong(hwndDlg, DWL_USER);
	SendMessage(GetDlgItem(hwndDlg, IDC_EDIT1), WM_SETTEXT, 0, (LPARAM)dtxview->text.c_str());
}

void GetListFiles(char *dir, vector<string> &files) {
	int ln;
	HANDLE hSearch;
	WIN32_FIND_DATA fd;
	char buf[4098];
	char buf0[1024];

	ZeroMemory(&fd, sizeof(WIN32_FIND_DATA));
	strcpy(buf, dir);
	strcat(buf, "\\*.dtx");
	hSearch = FindFirstFile(buf, &fd);
	if (hSearch == INVALID_HANDLE_VALUE)
		return;
	else {
		ln = strlen(fd.cFileName) - 4;
		strncpy(buf0, fd.cFileName, ln);
		buf0[ln] = '\0';
		files.push_back(buf0);
	}
	while (FindNextFile(hSearch, &fd)) {
		ln = strlen(fd.cFileName) - 4;
		strncpy(buf0, fd.cFileName, ln);
		buf0[ln] = '\0';
		files.push_back(buf0);
	}
	FindClose(hSearch);
}

void ReleaseItemData(HWND hwndOpenDoc) {
	string *s;
	int ln = SendDlgItemMessage(hwndOpenDoc, IDC_LB, LB_GETCOUNT, 0, 0);
	if (ln > 0) {
		for (int i = 0; i < ln; i++) {
			s = (string*)SendDlgItemMessage(hwndOpenDoc, IDC_LB, LB_GETITEMDATA, i, 0);
			delete s;
		}
	}
}

BOOL CALLBACK FindProc(HWND hwndFind, UINT message, WPARAM wParam, LPARAM lParam) {
	string find;
	CDtxView *dtxview = (CDtxView*)GetWindowLong(GetParent(hwndFind), DWL_USER);
	switch (message)
	{
	case WM_INITDIALOG:
		//RefreshOpenDoc(hwndOpenDoc);
		//CenterDialog(hwndOpenDoc);
		return TRUE;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDC_BTNFIRST:
			find = GetEditText(GetDlgItem(hwndFind, IDC_EDIT));
			if (find != "")
				SendMessage(GetParent(hwndFind), MSG_FINDFIRST, 0, (LPARAM)(&find));
			return TRUE;
		case IDC_BTNNEXT:
			find = GetEditText(GetDlgItem(hwndFind, IDC_EDIT));
			if (find != "")
				SendMessage(GetParent(hwndFind), MSG_FINDNEXT, 0, (LPARAM)(&find));
			return TRUE;
		case IDC_CLOSE:
			dtxview->findpos = 0;
			DestroyWindow(hwndFind);
			return TRUE;
		default:
			return FALSE;
		}
	case WM_CLOSE:
		dtxview->findpos = 0;
		DestroyWindow(hwndFind);
		return TRUE;
	default:
		return FALSE;
	}
}

void FindFirst(HWND hwndDlg, string *find) {
	CDtxView *dtxview = (CDtxView*)GetWindowLong(hwndDlg, DWL_USER);
	basic_string<char>::size_type pos;
	dtxview->findpos = 0;
	string text = GetEditText(GetDlgItem(hwndDlg, IDC_EDIT1));
	pos = text.find(*find, 0);
	if (pos != basic_string<char>::npos) {
		// нашли !!!!!!
		dtxview->findpos = pos;
		SendDlgItemMessage(hwndDlg, IDC_EDIT1, EM_SETSEL, pos, pos + find->size());
		SendDlgItemMessage(hwndDlg, IDC_EDIT1, EM_SCROLLCARET, 0, 0);
		SetFocus(GetDlgItem(hwndDlg, IDC_EDIT1));
	}
}

void FindNext(HWND hwndDlg, string *find) {
	CDtxView *dtxview = (CDtxView*)GetWindowLong(hwndDlg, DWL_USER);
	basic_string<char>::size_type pos;
	string text = GetEditText(GetDlgItem(hwndDlg, IDC_EDIT1));
	pos = text.find(*find, dtxview->findpos + 1);
	if (pos != basic_string<char>::npos) {
		// нашли !!!!!!
		dtxview->findpos = pos;
		SendDlgItemMessage(hwndDlg, IDC_EDIT1, EM_SETSEL, pos, pos + find->size());
		SendDlgItemMessage(hwndDlg, IDC_EDIT1, EM_SCROLLCARET, 0, 0);
		SetFocus(GetDlgItem(hwndDlg, IDC_EDIT1));
	}
	else
		FindFirst(hwndDlg, find);
}

void WordText(string& s)
{
	// Variables that will be used and re-used in our calls
	DISPPARAMS dpNoArgs = { NULL, NULL, 0, 0 };
	DISPPARAMS dpArgsProp;
	VARIANT vArgsProp[1];
	VARIANT vResult;
	OLECHAR FAR* szFunction;
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	BSTR bstrTemp;
	EXCEPINFO exInfo;

	// IDispatch pointers for Word's objects
	IDispatch* pDispDocs;		//Documents collection
	IDispatch* pDispDoc;		//Document object
	IDispatch* pDispParagraphs;	// Paragraphs collection
	IDispatch* pDispParagraph;	// Paragraph object
	IDispatch* pDispRange;
	IDispatch* pDispFont;
	IDispatch* pPageSetup;

	// DISPIDs
	DISPID dispid;        //Documents property of Application 

						  //if (CoInitialize(NULL) != S_OK)
						  //{
						  //	MessageBox(NULL, "Ошибка инициализации COM", "", MB_OK);
						  //	return;
						  //}

						  // Get the CLSID for Word's Application Object
	CLSID clsid;
	CLSIDFromProgID(L"Word.Application", &clsid);

	// Create an instance of the Word application and obtain the pointer
	// to the application's IUnknown interface
	IUnknown* pUnk;
	HRESULT hr = ::CoCreateInstance(clsid,
		NULL,
		CLSCTX_SERVER,
		IID_IUnknown,
		(void**)&pUnk);

	// Query IUnknown to retrieve a pointer to the IDispatch interface
	IDispatch* pDispApp;
	hr = pUnk->QueryInterface(IID_IDispatch, (void**)&pDispApp);

	// Get IDispatch* for the Documents collection object
	szFunction = OLESTR("Documents");
	hr = pDispApp->GetIDsOfNames(IID_NULL, &szFunction, 1,
		LOCALE_USER_DEFAULT, &dispid);
	hr = pDispApp->Invoke(dispid, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
		&dpNoArgs, &vResult, NULL, NULL);
	pDispDocs = vResult.pdispVal;

	// Invoke the Add method on the Documents collection object
	// to create a new document in Word
	// Note that the Add method can take up to 3 arguments, all of 
	// which are optional. You are not passing it any so you are 
	// using an empty DISPPARAMS structure
	szFunction = OLESTR("Add");
	hr = pDispDocs->GetIDsOfNames(IID_NULL, &szFunction, 1,
		LOCALE_USER_DEFAULT,
		&dispid);
	hr = pDispDocs->Invoke(dispid, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_METHOD,
		&dpNoArgs, &vResult, NULL, NULL);
	pDispDoc = vResult.pdispVal;



	szFunction = OLESTR("Paragraphs");
	hr = pDispDoc->GetIDsOfNames(IID_NULL, &szFunction, 1,
		LOCALE_USER_DEFAULT, &dispid);
	//if (hr != S_OK) ErrHandler(hr
	hr = pDispDoc->Invoke(dispid, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
		&dpNoArgs, &vResult, &exInfo, NULL);
	if (hr != S_OK) ErrHandler(hr, exInfo);
	pDispParagraphs = vResult.pdispVal;

	szFunction = OLESTR("Add");
	hr = pDispParagraphs->GetIDsOfNames(IID_NULL, &szFunction, 1,
		LOCALE_USER_DEFAULT,
		&dispid);
	hr = pDispParagraphs->Invoke(dispid, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_METHOD,
		&dpNoArgs, &vResult, &exInfo, NULL);
	if (hr != S_OK) ErrHandler(hr, exInfo);
	pDispParagraph = vResult.pdispVal;

	szFunction = OLESTR("Range");
	hr = pDispParagraph->GetIDsOfNames(IID_NULL, &szFunction, 1,
		LOCALE_USER_DEFAULT, &dispid);
	hr = pDispParagraph->Invoke(dispid, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
		&dpNoArgs, &vResult, &exInfo, NULL);
	if (hr != S_OK) ErrHandler(hr, exInfo);
	pDispRange = vResult.pdispVal;

	szFunction = OLESTR("Text");
	hr = pDispRange->GetIDsOfNames(IID_NULL, &szFunction, 1,
		LOCALE_USER_DEFAULT, &dispid);
	wchar_t *w = new wchar_t[s.length() + 1];
	w[s.length()] = L'\0';
	MultiByteToWideChar(CP_ACP, 0, s.c_str(), -1, w, s.length());
	bstrTemp = ::SysAllocString(w);
	delete w;
	vArgsProp[0].vt = VT_BSTR;
	vArgsProp[0].bstrVal = bstrTemp;
	dpArgsProp.cArgs = 1;
	dpArgsProp.cNamedArgs = 1;
	dpArgsProp.rgvarg = vArgsProp;
	dpArgsProp.rgdispidNamedArgs = &dispidNamed;
	hr = pDispRange->Invoke(dispid, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
		&dpArgsProp, &vResult, &exInfo, NULL);
	if (hr != S_OK) ErrHandler(hr, exInfo);
	::SysFreeString(bstrTemp);

	szFunction = OLESTR("Font");
	hr = pDispRange->GetIDsOfNames(IID_NULL, &szFunction, 1,
		LOCALE_USER_DEFAULT, &dispid);
	hr = pDispRange->Invoke(dispid, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
		&dpNoArgs, &vResult, &exInfo, NULL);
	if (hr != S_OK) ErrHandler(hr, exInfo);
	pDispFont = vResult.pdispVal;

	szFunction = OLESTR("Name");
	hr = pDispFont->GetIDsOfNames(IID_NULL, &szFunction, 1,
		LOCALE_USER_DEFAULT, &dispid);
	bstrTemp = ::SysAllocString(L"Courier New");
	vArgsProp[0].vt = VT_BSTR;
	vArgsProp[0].bstrVal = bstrTemp;
	dpArgsProp.cArgs = 1;
	dpArgsProp.cNamedArgs = 1;
	dpArgsProp.rgvarg = vArgsProp;
	dpArgsProp.rgdispidNamedArgs = &dispidNamed;
	hr = pDispFont->Invoke(dispid, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
		&dpArgsProp, &vResult, &exInfo, NULL);
	if (hr != S_OK) ErrHandler(hr, exInfo);
	::SysFreeString(bstrTemp);

	szFunction = OLESTR("Size");
	hr = pDispFont->GetIDsOfNames(IID_NULL, &szFunction, 1,
		LOCALE_USER_DEFAULT, &dispid);
	vArgsProp[0].vt = VT_R4;
	vArgsProp[0].fltVal = 6.0;
	dpArgsProp.cArgs = 1;
	dpArgsProp.cNamedArgs = 1;
	dpArgsProp.rgvarg = vArgsProp;
	dpArgsProp.rgdispidNamedArgs = &dispidNamed;
	hr = pDispFont->Invoke(dispid, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
		&dpArgsProp, &vResult, &exInfo, NULL);
	if (hr != S_OK) ErrHandler(hr, exInfo);

	szFunction = OLESTR("PageSetup");
	hr = pDispDoc->GetIDsOfNames(IID_NULL, &szFunction, 1,
		LOCALE_USER_DEFAULT, &dispid);
	hr = pDispDoc->Invoke(dispid, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
		&dpNoArgs, &vResult, &exInfo, NULL);
	if (hr != S_OK) ErrHandler(hr, exInfo);
	pPageSetup = vResult.pdispVal;

	szFunction = OLESTR("LeftMargin");
	hr = pPageSetup->GetIDsOfNames(IID_NULL, &szFunction, 1,
		LOCALE_USER_DEFAULT, &dispid);
	vArgsProp[0].vt = VT_R4;
	vArgsProp[0].fltVal = 43.0;
	dpArgsProp.cArgs = 1;
	dpArgsProp.cNamedArgs = 1;
	dpArgsProp.rgvarg = vArgsProp;
	dpArgsProp.rgdispidNamedArgs = &dispidNamed;
	hr = pPageSetup->Invoke(dispid, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
		&dpArgsProp, &vResult, &exInfo, NULL);
	if (hr != S_OK) ErrHandler(hr, exInfo);

	szFunction = OLESTR("Visible");
	hr = pDispApp->GetIDsOfNames(IID_NULL, &szFunction, 1,
		LOCALE_USER_DEFAULT, &dispid);
	vArgsProp[0].vt = VT_BOOL;
	vArgsProp[0].boolVal = VARIANT_TRUE;
	dpArgsProp.cArgs = 1;
	dpArgsProp.cNamedArgs = 1;
	dpArgsProp.rgvarg = vArgsProp;
	dpArgsProp.rgdispidNamedArgs = &dispidNamed;
	hr = pDispApp->Invoke(dispid, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
		&dpArgsProp, &vResult, NULL, NULL);

	pPageSetup->Release();
	pDispRange->Release();
	pDispParagraph->Release();
	pDispParagraphs->Release();
	pDispDoc->Release();
	pDispDocs->Release();
	pDispApp->Release();
	pUnk->Release();
	//CoUninitialize();
}

void ErrHandler(HRESULT hr, EXCEPINFO excep)
{
	if (hr == DISP_E_EXCEPTION)
	{
		char errDesc[512];
		char errMsg[512];
		wcstombs(errDesc, excep.bstrDescription, 512);
		sprintf(errMsg, "Run-time error %d:\n\n %s",
			excep.scode & 0x0000FFFF,  //Lower 16-bits of SCODE
			errDesc);                  //Text error description
		::MessageBox(NULL, errMsg, "Server Error", MB_SETFOREGROUND |
			MB_OK);
	}
	else
	{
		LPVOID lpMsgBuf;
		::FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER |
			FORMAT_MESSAGE_FROM_SYSTEM |
			FORMAT_MESSAGE_IGNORE_INSERTS, NULL, hr,
			MAKELANGID(LANG_NEUTRAL,
				SUBLANG_DEFAULT), (LPTSTR)&lpMsgBuf,
			0, NULL);
		::MessageBox(NULL, (LPCTSTR)lpMsgBuf, "COM Error",
			MB_OK | MB_SETFOREGROUND);
		::LocalFree(lpMsgBuf);
	}

}
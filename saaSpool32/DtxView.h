#pragma once
#include "stdafx.h"
#include "Commctrl.h"
#include "dtx.h"
#include "option.h"

class CDtxView
{
public:
	CDtxView(HINSTANCE hInst);
	~CDtxView(void);
	void Show(char* fname);
	void Show(void);
	CDtx* dtx;
	std::string text;
	std::string CDtxView::readFile(const std::string& fileName);
	bool CDtxView::writeFile(const std::string& fileName);
private:
	void _Show();
	
public:
	HFONT GetFontForEditBox(void);
	HIMAGELIST hImageList;
	HFONT hEditFont;
	HWND hwndTB;
	COption* option;
	int CurSel;
	BOOL IsChanged;
	HWND hwndFind;
	int findpos;
	int page1;
	int page2;
	HINSTANCE hInstance;
};

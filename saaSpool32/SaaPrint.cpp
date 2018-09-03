#pragma warning(disable : 4996)
#include "StdAfx.h"
#include "SaaPrint.h"
#include <sstream>


CSaaPrint::CSaaPrint(string printerName)
	: lborder(10)
	, rborder(15)
	, tborder(10)
	, bborder(15)
	, MaximalFont(14)
	, MinimalFont(6)
{
	this->printerName = printerName;
}


CSaaPrint::~CSaaPrint(void)
{
}


bool CSaaPrint::PrintText(string& text, bool isPageNum)
{
	bool rt = false;
	HANDLE hPrinter = NULL;
	HDC hdc = NULL;
	DOCINFO di;
	int job = 0;
	LPDEVMODE devmode = NULL;
	string longest;
	//string text;
	try
	{
		//text = DOS2Win(text0);
		// определим самую длинную строку
		longest = getlongest(text);
		if (!OpenPrinter((LPSTR)printerName.c_str(), &hPrinter, NULL))
			throw;
		// получить размер структуры DEVMODE
		LONG sizeDevmode = DocumentProperties(NULL, hPrinter, NULL, devmode, devmode, 0);
		devmode = (LPDEVMODE)new BYTE[sizeDevmode];
		// установить портретную ориентацию, размер бумаги A4
		SetDevmode(hPrinter, devmode, sizeDevmode, DMORIENT_PORTRAIT);
		// создать контекст
		hdc = CreateDC(NULL, printerName.c_str(), NULL, devmode);
		if (hdc == NULL) throw;
		// вычислить размер шрифта
		int validFont = ValidFont(hdc, devmode, longest);
		if (validFont == 0)
		{
			// если нельзя разместить самую длинную строку приемлемым шрифтом
			// то удаляем контекст, устанавливаем альбомную ориентацию и заново
			// создаем контекст
			DeleteDC(hdc);
			SetDevmode(hPrinter, devmode, sizeDevmode, DMORIENT_LANDSCAPE);
			hdc = CreateDC(NULL, printerName.c_str(), NULL, devmode);
			validFont = ValidFont(hdc, devmode, longest);
			if (validFont == 0)
			{
				// распечатать текст приемлемым шрифтом не удалось ((((
				throw string("Не удалось распечатать текст приемлемым шрифтом");
			}
		}

		di.cbSize = sizeof(di);
		di.lpszDocName = _T("saaText");
		di.lpszOutput = NULL;
		di.lpszDatatype = NULL;
		di.fwType = 0;
		job = StartDoc(hdc, &di);
		if (job <= 0) throw;
		TEXTMETRIC tm;
		GetTextMetrics(hdc, &tm);
		int curPage = 1;
		siterator it = make_split_iterator(text, token_finder(boost::algorithm::is_any_of("\r\n")));
		while (it != siterator())
		{
			StartPage(hdc);
			it = AddContent2Page(hdc, tm, devmode, curPage, it, isPageNum);
			EndPage(hdc);
			curPage++;
		}
		rt = true;
	}
	catch (string e)
	{
		MessageBox(NULL, e.c_str(), "", MB_OK);
	}
	catch (...)
	{
		MessageBox(NULL, "Распечатать не удалось", "", MB_OK);
	}
	if (job > 0)
	{
		EndDoc(hdc);
		job = 0;
	}
	if (hdc != NULL)
	{
		DeleteDC(hdc);
		hdc = NULL;
	}
	if (devmode != NULL)
	{
		delete devmode;
		devmode = NULL;
	}
	if (hPrinter != NULL)
	{
		ClosePrinter(hPrinter);
		hPrinter = NULL;
	}
	return rt;
}


siterator CSaaPrint::AddContent2Page(HDC hdc, TEXTMETRIC tm, LPDEVMODE devmode, int nPage,
	siterator it, bool isPageNum)
{
	//char qq[1024];
	//char qq1[1024];
	//itoa(devmode->dmPrintQuality, qq, 10);
	//itoa(devmode->dmYResolution, qq1, 10);
	//MessageBox(NULL, qq, "", MB_OK);
	//MessageBox(NULL, qq1, "", MB_OK);
	int width, height;
	short yResolution;
	if (devmode->dmYResolution != 0)
		yResolution = devmode->dmYResolution;
	else
		yResolution = devmode->dmPrintQuality;
	int x = lborder * devmode->dmPrintQuality / 25.4;
	int y = tborder * yResolution / 25.4;
	int xcenter;
	int heightLine = tm.tmHeight / 2;// + tm.tmExternalLeading / 2;
	int y0;
	siterator It;
	if (devmode->dmOrientation == DMORIENT_PORTRAIT)
	{
		width = 210; height = 297;
	}
	else
	{
		width = 297; height = 210;
	}
	y0 = (height - bborder) * yResolution / 25.4;

	// оставляем место под номер страницы
	if (isPageNum)
		y0 -= heightLine * 2;
	string s;
	wstring wtxt;
	for (It = it; It != siterator() && y <= y0; ++It, y += heightLine)
	{
		s = boost::copy_range<std::string>(*It);
		wtxt = string2wstring(s, std::locale("russian_russia.1251"));
		TextOutW(hdc, x, y, wtxt.c_str(), wtxt.size());
	}
	// либо текст закончился, либо конец страницы
	char buf[10];
	if (isPageNum)
	{
		wsprintf(buf, "%d", nPage);
		std::string ss0("- ");
		ss0 += buf;
		ss0 += " -";
		xcenter = (lborder + (width - lborder - rborder) / 2) * devmode->dmPrintQuality / 25.4;
		SetTextAlign(hdc, TA_TOP | TA_CENTER);
		TextOutA(hdc, xcenter, y0 += heightLine * 2, ss0.c_str(), ss0.size());
		SetTextAlign(hdc, TA_TOP | TA_LEFT);
	}


	return ++It;
}


void CSaaPrint::SetDevmode(HANDLE hPrinter, LPDEVMODE devmode,
	LONG sizeDevmode, int orientation)
{
	memset(devmode, 0, sizeDevmode);
	LONG lReturn = DocumentProperties(NULL, hPrinter, NULL, devmode, NULL, DM_OUT_BUFFER);
	if (lReturn < 0) throw;
	devmode->dmPaperSize = DMPAPER_A4;
	devmode->dmOrientation = orientation;
	lReturn = DocumentProperties(NULL, hPrinter, NULL, devmode, devmode, DM_IN_BUFFER | DM_OUT_BUFFER);
	if (lReturn < 0) throw;
}

/*
Установить ориентацию = портрет

*/

int CSaaPrint::ValidFont(HDC hdc, LPDEVMODE devmode, const string& longest)
{
	HFONT font = NULL;
	SIZE sz;
	sz.cx = INT_MAX;
	int width;
	if (devmode->dmOrientation == DMORIENT_PORTRAIT)
		width = 210;
	else
		width = 297;
	// ширина в пикселах области печати
	int x = devmode->dmPrintQuality * (width - lborder - rborder) / 25.4;
	//std::string temp = this->wstring2string(longest);
	int pt = MaximalFont;
	while (sz.cx > x && pt >= MinimalFont)
	{
		if (font != NULL)
			DeleteObject(font);
		font = SetFont(hdc, pt, devmode);
		GetTextExtentPoint32A(hdc, longest.c_str(), longest.size(), &sz);
		pt--;
	}
	if (pt < MinimalFont)
		return 0;
	else
		return pt + 1;
}

HFONT CSaaPrint::SetFont(HDC hdc, int pt, LPDEVMODE devmode)
{
	HFONT textFont;
	textFont = CreateFont(
		pt * devmode->dmPrintQuality / 72,
		0,
		0,
		0,
		0,
		FALSE,
		FALSE,
		FALSE,
		DEFAULT_CHARSET,
		OUT_OUTLINE_PRECIS,
		CLIP_DEFAULT_PRECIS,
		DEFAULT_QUALITY,
		DEFAULT_PITCH,
		"Courier New");
	if (NULL != textFont)
	{
		SelectObject(hdc, textFont);
	}
	return textFont;
}

std::string CSaaPrint::getlongest(const std::string& in)
{
	string temp;
	string rt;
	int size;
	typedef boost::split_iterator<std::string::const_iterator> siterator;
	int len = -1;
	size_t endPos, startPos;
	for (siterator It =
		make_split_iterator(in, token_finder(boost::algorithm::is_any_of("\r\n")));
		It != siterator();
		++It)
	{
		temp = boost::copy_range<std::string>(*It);
		startPos = 0;
		endPos = temp.find_last_not_of(" ");
		temp = temp.substr(startPos, endPos - startPos + 1);
		size = temp.size();
		if (size > len)
		{
			len = temp.length();
			rt = temp;
		}
	}
	return rt;
}

std::string CSaaPrint::wstring2string(const std::wstring& in, std::locale loc)
{
	std::string out(in.length(), 0);
	std::wstring::const_iterator i = in.begin(), ie = in.end();
	std::string::iterator j = out.begin();

	for (; i != ie; ++i, ++j)
		*j = std::use_facet< std::ctype< wchar_t > >(loc).widen(*i);

	return out;
}

std::wstring CSaaPrint::string2wstring(const std::string& in, std::locale loc)
{
	std::wstring out(in.length(), 0);
	std::string::const_iterator i = in.begin(), ie = in.end();
	std::wstring::iterator j = out.begin();

	for (; i != ie; ++i, ++j)
		*j = std::use_facet< std::ctype< wchar_t > >(loc).widen(*i);

	return out;
}

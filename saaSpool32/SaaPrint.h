#pragma once
#include "winspool.h"
#include "WTypes.h"
#include <string>
#include <boost/algorithm/string.hpp>
#include <boost/algorithm/string/find_iterator.hpp>
#include <boost/algorithm/string/finder.hpp>
#include <boost/algorithm/string/iter_find.hpp>

using namespace std;

#pragma once
typedef boost::algorithm::split_iterator<std::string::iterator> siterator;
class CSaaPrint
{
public:
	CSaaPrint(string printerName);
	~CSaaPrint(void);
	bool PrintText(string& txt, bool isPageNum = false);
	std::string wstring2string(const std::wstring& in, std::locale loc = std::locale());
private:
	int lborder;
	int rborder;
	int tborder;
	int bborder;
	string printerName;
	void SetDevmode(HANDLE hPrinter, LPDEVMODE devmode,
		LONG sizeDevmode, int orientation);
	int ValidFont(HDC hdc, LPDEVMODE devmode, const string& longest);
	std::string getlongest(const std::string& in);
	std::wstring string2wstring(const std::string& in, std::locale loc = std::locale());
	HFONT SetFont(HDC hdc, int pt, LPDEVMODE devmode);
	int const MinimalFont;
	int const MaximalFont;
	siterator AddContent2Page(HDC hdc, TEXTMETRIC tm, LPDEVMODE devmode, int nPage,
		siterator it, bool isPageNum = false);
};


#include "StdAfx.h"
#include "saaCUtils.h"

int RTrimSpaces(char *s) {
	int i;
	if (!s)
		return 0;
	i = strlen(s);
	if (i == 0)
		return 0;
	i--;
	while (i >= 0) {
		if (s[i] != 0x20)
			return strlen(s) - i - 1;
		i--;
	}
	return strlen(s);
}

string GetEditText(HWND hwnd) {
	int ln = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0);
	char *s = new char[ln + 1];
	SendMessage(hwnd, WM_GETTEXT, ln + 1, (LPARAM)s);
	string rt(s);
	delete s;
	return rt;
}

void GetOnlyFilenameFromFullPath(string path, string& fname) {
	basic_string<char>::size_type pos, pos1;
	pos = path.find_last_of('\\');
	if (pos != basic_string<char>::npos)
		pos++;
	else
		pos = 0;
	pos1 = path.find_last_of('.');
	fname = path.substr(pos, pos1 - pos);
}

void GetTempFileName(string& s) {
	DWORD lnpath;
	char *buf;
	char fname[MAX_PATH];
	lnpath = GetTempPath(0, NULL);
	buf = new char[lnpath];
	GetTempPath(lnpath + 1, buf);
	GetTempFileName(buf, "saa", 0, fname);
	delete buf;
	s = string(fname);
}
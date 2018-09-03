#include "StdAfx.h"
#include ".\option.h"

#pragma warning(disable : 4996)


COption::COption(void)
{
	string df;

	if (GetDefaultConfig(df)) {
		this->Load((char*)df.c_str());
	}
	else {
		this->doclist = NULL;
		this->dtxdir = NULL;
		this->name = NULL;
	}
}

COption::~COption(void)
{
	if (this->name)
		delete this->name;
	if (this->doclist)
		delete this->doclist;
	if (this->dtxdir)
		delete this->dtxdir;
}

COption::COption(const char* parname, const char* pardoclist, const char* pardtxdir)
{
	this->name = new char[strlen(parname) + 1];
	strcpy(this->name, parname);
	this->doclist = new char[strlen(pardoclist) + 1];
	strcpy(this->doclist, pardoclist);
	this->dtxdir = new char[strlen(pardtxdir) + 1];
	strcpy(this->dtxdir, pardtxdir);
}

void COption::Save()
{
	HKEY hKey;
	char s[100];
	strcpy(s, "Software\\saaSpool\\Options\\");
	strcat(s, this->name);
	if (RegCreateKeyEx(HKEY_CURRENT_USER, s, 0, 0, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, &hKey, 0) == ERROR_SUCCESS) {
		RegSetValueEx(hKey, "doclist", 0, REG_SZ, (const BYTE*)this->doclist, strlen(this->doclist) + 1);
		RegSetValueEx(hKey, "dtxdir", 0, REG_SZ, (const BYTE*)this->dtxdir, strlen(this->dtxdir) + 1);
		RegCloseKey(hKey);
	}
}

COption::COption(char* parname)
{
	// parname указатель на буфер размеров 100 байт, 1-ый байт \0
	// это название конфигурации
	this->Load(parname);
}

BOOL COption::Load(char *df) {
	HKEY hKey;
	DWORD type = REG_SZ;
	DWORD size;
	char s[100];
	strcpy(s, "Software\\saaSpool\\Options\\");
	strcat(s, df);
	this->name = new char[strlen(df) + 1];
	strcpy(this->name, df);
	if (RegOpenKeyEx(HKEY_CURRENT_USER, s, 0, KEY_ALL_ACCESS, &hKey) == ERROR_SUCCESS) {
		RegQueryValueEx(hKey, "doclist", 0, &type, NULL, &size);
		this->doclist = new char[size];
		RegQueryValueEx(hKey, "doclist", 0, &type, (LPBYTE)this->doclist, &size);
		RegQueryValueEx(hKey, "dtxdir", 0, &type, NULL, &size);
		this->dtxdir = new char[size];
		RegQueryValueEx(hKey, "dtxdir", 0, &type, (LPBYTE)this->dtxdir, &size);
		RegCloseKey(hKey);
		return TRUE;
	}
	else
	{
		this->doclist = NULL;
		this->dtxdir = NULL;
		this->name = NULL;
		return FALSE;
	}
}

BOOL COption::GetDefaultConfig(string& df)
{
	HKEY hKey;
	DWORD type = REG_SZ;
	DWORD size = 0;
	char *s;
	if (RegOpenKeyEx(HKEY_CURRENT_USER, "Software\\saaSpool", 0, KEY_ALL_ACCESS, &hKey) == ERROR_SUCCESS) {
		if (RegQueryValueEx(hKey, "curconfig", 0, &type, NULL, &size) != ERROR_SUCCESS)
		{
			return FALSE;
		}
		if (size == 1)
			return FALSE;
		s = new char[size];
		RegQueryValueEx(hKey, "curconfig", 0, &type, (LPBYTE)s, &size);
		df.assign(s);
		delete s;
		RegCloseKey(hKey);
		return TRUE;
	}
	else
		return FALSE;
}

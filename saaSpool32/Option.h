#pragma once
#include <string>

using namespace std;

class COption
{
public:
	COption(void);
	~COption(void);
	char* name;
	char* doclist;
	char* dtxdir;
	COption(const char* parname, const char* pardoclist, const char* pardtxdir);
	void Save(void);
	BOOL Load(char *df);
	COption(char* parname);
	static BOOL GetDefaultConfig(string& df);
};

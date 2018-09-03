#pragma once
#include "dbf.h"
#include <string>
#include <vector>
//#include "PairInt.h"

using namespace std;

struct DTX_STRUC {
	int page;
	char name[200];
	char *text;
};

class CDtx
{
public:
	CDtx(void);
	~CDtx(void);
	// Открывает DTX файл
	bool Open(char* fname);
private:
	char fname[200];
	char headname[4098];
	DTX_STRUC *dtx;

public:
	int pagenumber;
	void Close(void);
	char* GetText(int page);
	char* GetPageHeader(int page);
	void PrintPage(string pName, int page);
	int compress;
	int copies;
private:
	BOOL GetCC(char* fname, int* _compress, int* _copies, int* _lmargin);
	void PrintText(string pName, string& s);
public:
	int mlen;
	void PrintAll(string pName, bool bNewPage);
	void WordAll();
	void WordText(string& s);
	void PrintFromPage(string pName, int page, bool bNewPage);
	void Save(void);
private:
	int Update_mlen(void);
public:
	BOOL IsChanged;
	void SavePage(const char *s, int page);
private:
	int dbffieldcount;
	DBF_STRUCT *dbfstruct;
public:
	void PrintSelectedPages(string pName, bool bNewPage, int* selitems, int selcount);
	void WordSelectedPages(int* selitems, int selcount);
	int lmargin;
	void CreateNew(char* docname, char* filename);
	void SetPageText(int pagenumber, char* text);
	int AddPage(char* pagename);
private:
	void SetMaxLen(string text);
public:
	void SplitPages(void);
	vector<string> vPages;
private:
	const static int maxPortret = 157;
public:
	void PrintNumPages(string pName, int page1, int page2);
private:
	string GetTextStr(int nstr);
public:
	string GetTextString(int nstr);
	int GetMaxPortret(void);
	string GetHead(void);
};

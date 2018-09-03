#include "StdAfx.h"
#include "math.h"
#include "OaIdl.h"
#include ".\dtx.h"
#include "dbf.h"
#include <vector>
#include "saaCUtils.h"
#include "CSpooler.h"
#include <string>
#include "Option.h"
#include "SaaPrint.h"

#define MAXPAGES 1000
#pragma warning(disable : 4996)


double round(double number, int pos);
char* mystrcat(char* dest, char* src);
void padl(char* s, int n);
void ErrHandler(HRESULT hr, EXCEPINFO excep);

using namespace std;

CDtx::CDtx(void)
	: compress(0)
	, copies(0)
	, mlen(0)
	, IsChanged(FALSE)
	, dbfstruct(NULL)
	, dbffieldcount(0)
	, lmargin(0)
{
}

CDtx::~CDtx(void)
{
}

// ќткрывает DTX файл
bool CDtx::Open(char* _fname)
{
	CDbf dbf;
	int n = 0;
	int page;
	vector<int> v;
	int _compress, _copies, _lmargin;
	string textbuf;
	int sz;
	char *t;

	_strupr(_fname);
	// пытаемс€ получить сведени€ о шрифте и кол-ве копий из doclist.dbf
	if (GetCC(_fname, &_compress, &_copies, &_lmargin)) {
		this->compress = _compress;
		this->copies = _copies;
		this->lmargin = _lmargin;
	}
	else {
		this->compress = 0;
		this->copies = 1;
		this->lmargin = 0;
	}

	// запоминаем им€ файла
	strcpy(this->fname, _fname);
	if (dbf.Open(fname) == 0)
	{
		//MessageBox(NULL, "облом", "", MB_OK);
		return false;
	}
	this->dbffieldcount = dbf.FieldCount();
	this->dbfstruct = new DBF_STRUCT[this->dbffieldcount];
	memcpy(this->dbfstruct, dbf.dbfstruct, this->dbffieldcount * sizeof(DBF_STRUCT));


	// получаем длину пол€ TEXTLINE и резервируем под нее буфер textline
	// с учетом \r\n
	int nTextLine = dbf.FieldLen("TEXTLINE") + 2;

	// считаем количество страниц pagenumber и число строк на каждой странице
	dbf.GoTop();
	pagenumber = 0;
	while (dbf.Eof() != 1) {
		n = dbf.FieldGetInt(string("TEXTSIGN"));
		if (n == 1 || n == 2) {
			v.push_back(0);
			pagenumber++;	// увеличиваем счетчик страниц
		}
		if (n == 0) {
			if (v.size() > 0)
				v.back()++;		// число строк увеличиваем на единицу
			else
			{
				delete this->dbfstruct;
				return false;
			}
		}
		dbf.Skip(1);
	}

	dtx = new DTX_STRUC[pagenumber]; // заказываем пам€ть под структуру документа

	mlen = -1;
	page = 0;
	dbf.GoTop();
	while (dbf.Eof() != 1) {			// еще один проход по документу
		n = dbf.FieldGetInt(string("TEXTSIGN"));
		if (n == 1 || n == 2) {				// заголовок страницы
			if (n == 2) {				// заголовок файла
				dbf.FieldGetString(string("TEXTLINE"), textbuf);
				strcpy(headname, textbuf.c_str());
			}
			dtx[page].page = page;
			dbf.FieldGetString(string("TEXTLINE"), textbuf);
			strcpy(dtx[page].name, textbuf.c_str());
			dtx[page].text = new char[v[page] * nTextLine + 1];
			dtx[page].text[0] = '\0';
			t = dtx[page].text;
			page++;
		}
		if (n == 0) {				// строка текста
			dbf.FieldGetString(string("TEXTLINE"), textbuf);
			sz = textbuf.size();//strlen(textbuf.c_str());
			if (sz > mlen)
				mlen = sz;
			t = mystrcat(t, (char*)textbuf.c_str());
			t = strcat(t, "\r\n");
		}

		dbf.Skip(1);
	}
	dbf.Close();
	return true;
}

void CDtx::Close(void)
{
	int i;
	for (i = 0; i < pagenumber; i++) {
		delete dtx[i].text;
	}
	if (this->dbfstruct)
		delete this->dbfstruct;
	vPages.clear();
}

char* CDtx::GetText(int page)
{
	if (page < this->pagenumber) {
		char *s = new char[strlen(dtx[page].text) + 1];
		strcpy(s, dtx[page].text);
		return s;
	}
	else
		return NULL;
}

char* CDtx::GetPageHeader(int page)
{
	if (page < this->pagenumber) {
		char *s = new char[strlen(dtx[page].name) + 1];
		strcpy(s, dtx[page].name);
		return s;
	}
	else
		return NULL;
}

void CDtx::PrintPage(string pName, int page)
{
	string s;
	char *s0 = this->GetText(page);
	if (s0) {
		if (s0[0] != '\0')
		{
			s = string(s0);
			this->PrintText(pName, s);
			delete s0;
		}
	}
}

BOOL CDtx::GetCC(char* fname, int* _compress, int* _copies, int* _lmargin)
{
	COption opt;
	string shortfname;
	string _fname;
	CDbf _doclist;
	if (opt.name == NULL)
		return FALSE;
	GetOnlyFilenameFromFullPath(string(fname), shortfname);
	if (!_doclist.Open(opt.doclist))
		return FALSE;
	while (_doclist.Eof() != 1) {
		_doclist.FieldGetString(string("FNAME"), _fname);
		if (_fname == string(shortfname)) {
			*_compress = _doclist.FieldGetInt(string("COMPRESS"));
			*_copies = _doclist.FieldGetInt(string("COPIES"));
			*_lmargin = _doclist.FieldGetInt(string("OFFSET"));
			_doclist.Close();
			return TRUE;
		}
		_doclist.Skip(1);
	}
	_doclist.Close();
	return FALSE;;
}


double round(double number, int pos) // pos Ч число знаков после зап€той, если pos<0, то число знаков до зап€той.
{
	double base = 10., coeff = 1.;
	int i;

	for (i = 0; i<pos; i++) coeff *= base;

	number *= coeff;
	number = floor(number + 0.5);
	number /= coeff;

	return number;
}

void CDtx::PrintAll(string pName, bool bNewPage)
{
	string s("");
	char *s0 = NULL;
	for (int i = 0; i < this->pagenumber; i++) {
		s0 = this->GetText(i);
		if (s0) {
			if (s0[0] != '\0')
			{
				if (bNewPage)
					this->PrintText(pName, string(s0));
				else
					s.append(string(s0));
			}
			delete s0;
		}
	}
	if (!bNewPage)
		this->PrintText(pName, s);
}

void CDtx::PrintFromPage(string pName, int page, bool bNewPage)
{
	string s("");
	char *s0 = NULL;
	for (int i = page; i < this->pagenumber; i++) {
		s0 = this->GetText(i);
		if (s0) {
			if (s0[0] != '\0')
			{
				if (bNewPage)
					this->PrintText(pName, string(s0));
				else
					s.append(string(s0));
			}
			delete s0;
		}
	}
	if (!bNewPage)
		this->PrintText(pName, s);
}


void CDtx::WordAll()
{
	string s("");
	char *s0 = NULL;
	for (int i = 0; i < this->pagenumber; i++) {
		s0 = this->GetText(i);
		if (s0) {
			s.append(string(s0));
			delete s0;
		}
	}
	this->WordText(s);
}

void CDtx::WordText(string& s)
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
						  //	MessageBox(NULL, "ќшибка инициализации COM", "", MB_OK);
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

void CDtx::PrintText(string pName, string& s)
{
	CSaaPrint print(pName);
	print.PrintText(s, false);
}

void CDtx::Save(void)
{
	int page;
	CDbf dbf;
	string tfname;
	string s;
	basic_string<char>::size_type pos, pos0;
	if (this->IsChanged) {
		GetTempFileName(tfname);
		this->mlen = Update_mlen();
		dbfstruct[1].len = this->mlen;
		dbf.Create((char*)tfname.c_str(), dbfstruct, this->dbffieldcount);
		if (dbf.Open((char*)tfname.c_str()) == 0) {
			MessageBox(NULL, "ќшибка при сохранении", "", MB_OK);
			return;
		}
		dbf.Append();
		dbf.FieldPutInt(string("TEXTSIGN"), 2);
		dbf.FieldPutString(string("TEXTLINE"), this->headname);
		for (page = 0; page < this->pagenumber; page++) {
			dbf.Append();
			dbf.FieldPutInt(string("TEXTSIGN"), 1);
			dbf.FieldPutString(string("TEXTLINE"), this->dtx[page].name);
			s = string(dtx[page].text);
			pos0 = 0;
			do {
				pos = s.find("\r\n", pos0);
				if (pos == basic_string<char>::npos) {
					pos = s.substr(pos0).size();
					if (pos != 0) {
						dbf.Append();
						dbf.FieldPutInt(string("TEXTSIGN"), 0);
						dbf.FieldPutString(string("TEXTLINE"), s.substr(pos0));
					}
					break;
				}
				dbf.Append();
				dbf.FieldPutInt(string("TEXTSIGN"), 0);
				dbf.FieldPutString(string("TEXTLINE"), s.substr(pos0, pos - pos0));
				pos0 = pos + 2;
			} while (true);
		}
		dbf.Close();
		if (!CopyFile((char*)tfname.c_str(), this->fname, FALSE)) {
			MessageBox(NULL, "ќшибка при сохранении", "", MB_OK);
			return;
		}
		DeleteFile((char*)tfname.c_str());
		this->IsChanged = FALSE;
	}
}

int CDtx::Update_mlen(void)
{
	int i;
	string s;
	basic_string<char>::size_type pos, pos0, _mlen;
	_mlen = 0;
	for (i = 0; i < this->pagenumber; i++) {
		s = string(dtx[i].text);
		pos0 = 0;
		do {
			pos = s.find("\r\n", pos0);
			if (pos == basic_string<char>::npos) {
				pos = s.substr(pos0).size();
				if (pos > _mlen)
					_mlen = pos;
				break;
			}
			if ((pos - pos0) > _mlen)
				_mlen = pos - pos0;
			pos0 = pos + 2;
		} while (true);
	}
	return _mlen;
}

void CDtx::SavePage(const char *s, int page)
{
	delete dtx[page].text;
	dtx[page].text = new char[strlen(s) + 1];
	strcpy(dtx[page].text, s);
}

void CDtx::PrintSelectedPages(string pName, bool bNewPage, int* selitems, int selcount)
{
	string s("");
	char *s0 = NULL;
	for (int i = 0; i < selcount; i++) {
		s0 = this->GetText(selitems[i]);
		if (s0) {
			if (s0[0] != '\0')
			{
				if (bNewPage)
					this->PrintText(pName, string(s0));
				else
					s.append(string(s0));
			}
			delete s0;
		}
	}
	if (!bNewPage)
		this->PrintText(pName, s);
}

void CDtx::WordSelectedPages(int* selitems, int selcount)
{
	string s("");
	char *s0 = NULL;
	for (int i = 0; i < selcount; i++) {
		s0 = this->GetText(selitems[i]);
		if (s0) {
			s.append(string(s0));
			delete s0;
		}
	}
	this->WordText(s);
}

void CDtx::CreateNew(char* docname, char* filename)
{
	this->compress = 0;
	this->copies = 1;
	this->lmargin = 0;
	this->dbffieldcount = 8;
	this->dbfstruct = new DBF_STRUCT[this->dbffieldcount];
	ZeroMemory(this->dbfstruct, sizeof(DBF_STRUCT) * this->dbffieldcount);

	strcpy(this->dbfstruct[0].dbf_name, "TEXTSIGN");
	strcpy(this->dbfstruct[0].dbf_type, "N");
	this->dbfstruct[0].len = 2;
	this->dbfstruct[0].dec = 0;

	strcpy(this->dbfstruct[1].dbf_name, "TEXTLINE");
	strcpy(this->dbfstruct[1].dbf_type, "C");
	this->dbfstruct[1].len = 80;  // ??????????????????????????????????????????????????
	this->dbfstruct[1].dec = 0;

	strcpy(this->dbfstruct[2].dbf_name, "ENHBEGIN");
	strcpy(this->dbfstruct[2].dbf_type, "N");
	this->dbfstruct[2].len = 3;
	this->dbfstruct[2].dec = 0;

	strcpy(this->dbfstruct[3].dbf_name, "ENHEND");
	strcpy(this->dbfstruct[3].dbf_type, "N");
	this->dbfstruct[3].len = 3;
	this->dbfstruct[3].dec = 0;

	strcpy(this->dbfstruct[4].dbf_name, "ENHCODE");
	strcpy(this->dbfstruct[4].dbf_type, "N");
	this->dbfstruct[4].len = 1;
	this->dbfstruct[4].dec = 0;

	strcpy(this->dbfstruct[5].dbf_name, "ENHBEGSTR");
	strcpy(this->dbfstruct[5].dbf_type, "C");
	this->dbfstruct[5].len = 30;
	this->dbfstruct[5].dec = 0;

	strcpy(this->dbfstruct[6].dbf_name, "ENHENDSTR");
	strcpy(this->dbfstruct[6].dbf_type, "C");
	this->dbfstruct[6].len = 30;
	this->dbfstruct[6].dec = 0;

	strcpy(this->dbfstruct[7].dbf_name, "ENHCODESTR");
	strcpy(this->dbfstruct[7].dbf_type, "C");
	this->dbfstruct[7].len = 30;
	this->dbfstruct[7].dec = 0;

	_strupr(filename);
	strcpy(this->fname, filename);
	strcpy(this->headname, docname);
	this->IsChanged = FALSE;
	this->mlen = 0;
	this->pagenumber = 0;

	//dtx = new DTX_STRUC[MAXPAGES];
}

int CDtx::AddPage(char* pagename)
{
	DTX_STRUC *dtx0;
	if (!dtx) {
		dtx = new DTX_STRUC[1];
		ZeroMemory(dtx, sizeof(DTX_STRUC));
		strcpy(dtx->name, pagename);
		dtx->page = 0;
		this->pagenumber = 1;
		this->IsChanged = TRUE;
		return 0;
	}
	else {
		dtx0 = new DTX_STRUC[this->pagenumber + 1];
		this->pagenumber++;
		this->IsChanged = TRUE;
		ZeroMemory(dtx0, sizeof(DTX_STRUC) * this->pagenumber);
		for (int i = 0; i < this->pagenumber - 1; i++) {
			strcpy(dtx0[i].name, dtx[i].name);
			dtx0[i].page = dtx[i].page;
			dtx0[i].text = new char[strlen(dtx[i].text) + 1];
			strcpy(dtx0[i].text, dtx[i].text);
			delete dtx[i].text;
		}
		strcpy(dtx0[pagenumber - 1].name, pagename);
		dtx0[pagenumber - 1].page = pagenumber - 1;
		delete dtx;
		dtx = dtx0;
		return pagenumber - 1;
	}
}

void CDtx::SetPageText(int pagenumber, char* text)
{
	if (dtx[pagenumber].text)
		delete dtx[pagenumber].text;
	dtx[pagenumber].text = new char[strlen(text) + 1];
	strcpy(dtx[pagenumber].text, text);
	SetMaxLen(string(text));
}

void CDtx::SetMaxLen(string text)
{
	int sz;
	basic_string<char>::size_type pos0 = 0;
	basic_string<char>::size_type pos;
	pos = text.find(string("\r\n"), pos0);
	while (pos != basic_string<char>::npos) {
		// нашли !!!!!!
		sz = pos - pos0;
		if (sz > mlen) {
			mlen = sz;
			this->dbfstruct[1].len = mlen;
		}
		//pos0 += 2;
		pos0 = pos + 2;
		pos = text.find(string("\r\n"), pos0);
	}
	sz = text.size() - pos0;
	if (sz > mlen) {
		mlen = sz;
		this->dbfstruct[1].len = mlen;
	}
}

char* mystrcat(char* dest, char* src)
{
	while (*dest) dest++;
	while (*dest++ = *src++);
	return --dest;
}


void CDtx::SplitPages(void)
{
	int KolLine;
	int sz;
	int gline = 0;
	string text;
	char* lf = "\r\nCтpaницa     \r\n";
	char* lf1 = "\r\n\r\nCтpaницa     \r\n";
	int nstr = 0;
	char tbuf[1024];

	basic_string<char>::size_type pos0;
	basic_string<char>::size_type pos;
	//basic_string<char>::size_type pos1;	
	vPages.clear();
	vPages.push_back(string(""));
	int maxPortret0 = GetMaxPortret();
	if (this->mlen > maxPortret0)
		KolLine = 55;
	else
		KolLine = 82;

	//////////////////////////////////////////////
	for (int i = 0; i < this->pagenumber; i++) {
		text = string(dtx[i].text);
		pos0 = 0;
		pos = text.find(string("\r\n"), pos0);
		while (pos != basic_string<char>::npos) {
			// нашли !!!!!!
			gline++;
			if (gline == KolLine) {
				vPages[nstr].append(text.substr(pos0, pos + 2 - pos0));
				pos0 = pos + 2;
				text.insert(pos0, lf);
				_itoa(nstr + 1, tbuf, 10);
				padl(tbuf, 5);
				text.replace(pos0 + 10, 5, tbuf);
				vPages[nstr].append(text.substr(pos0, pos + 19 - pos0));
				pos0 += 17;
				gline = 0;
				nstr++;
				vPages.push_back(string(""));
			}
			else {
				vPages[nstr].append(text.substr(pos0, pos + 2 - pos0));
				pos0 = pos + 2;
			}
			pos = text.find(string("\r\n"), pos0);
		}
		//gline++;
		sz = text.size() - pos0;
		if (sz > 0) {  // есть последн€€ строка без \r\n
			gline++;
			if (gline == KolLine) {
				text.append(lf1);
				_itoa(nstr + 1, tbuf, 10);
				padl(tbuf, 5);
				text.replace(text.size() - 7, 5, tbuf);
				vPages[nstr].append(text.substr(pos0, sz + 19));
				pos0 += 19;
				gline = 0;
				nstr++;
			}
		}
		delete dtx[i].text;
		dtx[i].text = new char[text.size() + 1];
		strcpy(dtx[i].text, text.c_str());
	}
	/////////////////////////////////////////////
}

void padl(char* s, int n) {
	char buf[1024];
	int ln = strlen(s);
	if (ln == n)
		return;
	if (ln < n) {
		for (int i = 0; i < n - ln; i++)
			*(buf + i) = ' ';
		buf[n - ln] = '\0';
		strcat(buf, s);
		strcpy(s, buf);
		return;
	}
}
void CDtx::PrintNumPages(string pName, int page1, int page2)
{
	string s("");
	for (int i = page1 - 1; i < page2; i++) {
		s.append(vPages[i]);
	}
	PrintText(pName, s);
}

int CDtx::GetMaxPortret(void)
{
	HKEY hKey;
	DWORD data = 0;
	DWORD size = 4;
	DWORD type = REG_DWORD;
	int mp = 0;
	int rt = 0;
	if (RegOpenKeyEx(HKEY_CURRENT_USER, "Software\\saaSpool", 0, KEY_ALL_ACCESS, &hKey) == ERROR_SUCCESS) {
		if (RegQueryValueEx(hKey, "maxportret", 0, &type, (LPBYTE)&data, &size) != ERROR_SUCCESS)
			rt = this->maxPortret;
		else
		{
			rt = (int)data;
		}
		RegCloseKey(hKey);
		return rt;
	}
	else
		return this->maxPortret;
}

string CDtx::GetHead(void)
{
	return string(headname);
}

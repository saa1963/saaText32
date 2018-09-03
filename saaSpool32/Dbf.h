#pragma once
#include <fstream>
#include <string>
using namespace std;

struct DBF_STRUCT {
	char dbf_name[11];
	char dbf_type[2];
	int len;
	int dec;
};

class CDbf
{
public:
	DBF_STRUCT* dbfstruct;
	int Locate();
	char* m_LocateString;
	void Zap();
	void _FillBuffer(char *s);
	int FieldCount();
	char* NameOfField(int pos);
	int Pack();
	void ReCall();
	void Delete();
	int IsOpen();
	char FieldType(char* fieldname);
	int Eof();
	char* FieldGet(char *fieldname, char* value, int& lenbuf);
	void FieldPut(char* value, char* fieldname);
	void Skip(long step);
	void GoBottom();
	int Deleted();
	void GoTop();
	unsigned long Recno();
	int GoTo(unsigned long recno);
	int Append();
	int _GetCountRec();
	void Close();
	int Open(char *fname);
	static int Create(char *fname, DBF_STRUCT *d, int cntFields);
	CDbf();
	virtual ~CDbf();

protected:
	int LocateCalc();
	unsigned long base0;
	int IsInit;
	char* filename;
	int GetIndexField(char* s, int& pos);
	void Skip2(long step);
	void Skip1(long step);
	int PutBuffer();
	void ClearBOF_EOF();
	void SetEOF();
	int ClearBuffer();
	int GetBuffer();
	char* _buffer;
	int CountField;
	int _bof;
	int _eof;
	unsigned long lastrec;
	unsigned long _recno;
	int len_field;
	fstream f;

private:
	static void _putdate1(ofstream& f1);
public:
	int FieldLen(char* fieldname);
	void FieldGetString(string field, string& value);
	int FieldGetInt(string field);
	void FieldPutString(string field, string value);
	void FieldPutInt(string field, int value);
};
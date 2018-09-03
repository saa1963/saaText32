#include "stdafx.h"
#include "Dbf.h"
#include <memory.h>
#include <fstream>
#include <strstream>
#include <iomanip>
#include <time.h>
#include <stdlib.h>
#include <string.h>
#include ".\dbf.h"

using namespace std;
#pragma warning(disable : 4996)


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CDbf::CDbf()
{
	m_LocateString = NULL;
}

CDbf::~CDbf()
{
	if (f.is_open()) {
		f.close();
		delete[] dbfstruct;
		delete[] _buffer;
		delete[] filename;
	}
	delete m_LocateString;
}

int CDbf::Create(char *fname, DBF_STRUCT *d, int cntFields)
{
	ofstream f1;
	int i, j;
	short f;
	char temp[2];
	if (d && cntFields > 0) {
		f1.open(fname, ios::trunc | ios::binary);
		f1.put('\x3');
		_putdate1(f1);
		// число записей в .dbf
		f1.write("\x0\x0\x0\x0", 4);
		// положение первой записи с данными
		f = 32 + 32 * cntFields + 2;
		f1.write((char*)&f, 2);
		// длина записи с данными
		for (i = 0, f = 0; i < cntFields; i++)
			f += d[i].len;
		f++;
		f1.write((char*)&f, 2);
		// зарезервировано
		f1.write("\x0\x0", 2);
		// незавершенная транзакция (DBASE IV)
		f1.put('\x0');
		// шифрация данных (DBASE IV)
		f1.put('\x0');
		// free record thread (reserved for LAN only)
		f1.write("\x0\x0\x0\x0", 4);
		// (Reserved for multi-user dBASE) (DBASE III+ -)
		f1.write("\x0\x0\x0\x0\x0\x0\x0\x0", 8);
		// MDX flag (DBASE IV)
		f1.put('\x0');
		// 866 codepage (FOXPRO)
		//f1.put('\x66');
		f1.put('\x0');
		// зарезервировано
		f1.write("\x0\x0", 2);

		for (i = 0; i < cntFields; i++) {
			// Field name
			j = 0;
			_strupr(d[i].dbf_name);
			while (d[i].dbf_name[j]) {
				f1.put(d[i].dbf_name[j]);
				j++;
			}
			while (j < 11) {
				f1.put('\x0');
				j++;
			}
			// Field type
			temp[0] = d[i].dbf_type[0];
			temp[1] = '\0';
			_strupr(temp);
			d[i].dbf_type[0] = temp[0];
			d[i].dbf_type[1] = '\0';
			f1.put(d[i].dbf_type[0]);
			// Field data address
			f1.write("\x0\x0\x0\x0", 4);
			// Field length
			f1.put((char)d[i].len);
			// Field decimal
			f1.put((char)d[i].dec);
			// 
			f1.write("\x0\x0\x0\x0\x0\x0\x0\x0\x0\x0\x0\x0\x0\x0", 14);
		}
		// Terminator
		f1.put('\xd');
		//for (j = 0; i < 264; i++)
		//	f1.put('\x0');
		// Terminator
		f1.put('\x0');
		//f1.put('\x1a');
		f1.close();
		return 1;
	}
	else
		return 0;
}

int CDbf::Open(char *fname)
{
	short n;
	char n1;
	int i;
	unsigned char buf[32];
	wchar_t *fname0 = new wchar_t[strlen(fname) + 1];
	wmemset(fname0, 0, strlen(fname) + 1);
	MultiByteToWideChar(CP_ACP, 0, fname, -1, fname0, strlen(fname) + 1);
	f.open(fname0, ios_base::in | ios_base::out | ios_base::binary);
	delete fname0;
	if (f.good()) {
		IsInit = 1;

		// вычисляем количество полей в файле
		f.seekg(8, ios::beg);
		f.read((char*)&n, 2);
		base0 = n;
		do {
			n--;
			f.seekg(n, ios::beg);
			f.read((char*)&n1, 1);
		} while ((unsigned char)n1 != '\x0d');
		CountField = (n - 32) / 32;

		// создаем структуру dbfstruct и заполняем ее
		dbfstruct = new DBF_STRUCT[CountField];
		len_field = 1;
		for (i = 0; i < CountField; i++) {
			f.seekg(32 * (i + 1), ios::beg);
			f.read((char*)buf, sizeof(buf));
			memcpy(dbfstruct[i].dbf_name, buf, 11);
			memcpy(dbfstruct[i].dbf_type, &(buf[11]), 2);
			dbfstruct[i].len = buf[16];
			dbfstruct[i].dec = buf[17];
			len_field += dbfstruct[i].len;
		}

		// число записей в файле

		lastrec = _GetCountRec();
		_recno = 0;

		// выделяем буфер под запись
		_buffer = new char[len_field];

		// сохраняем имя файла
		filename = new char[strlen(fname) + 1];
		strcpy(filename, fname);

		GoTop();
		return 1;
	}
	else
		return 0;
}

void CDbf::Close()
{
	if (f.is_open()) {
		f.close();
		delete[] dbfstruct;
		delete[] _buffer;
		delete[] filename;
	}
}

int CDbf::Append()
{
	if (f.is_open()) {
		PutBuffer();
		ClearBuffer();
		_recno = lastrec + 1;
		lastrec++;
		PutBuffer();
		f.seekp(4, ios::beg);
		f.write((char*)(&lastrec), sizeof(lastrec));
		ClearBOF_EOF();
		return 1;
	}
	else
		return 0;
}

int CDbf::_GetCountRec() {
	unsigned long n;
	if (f.is_open()) {
		f.seekg(4, ios::beg);
		f.read((char*)&n, sizeof(n));
		return n;
	}
	else
		return -1;
}

// записывает текущий буфер на диск, если номер записи существует
// устанавливает его, заполняет буфер текущей записью и возвращает 1, 
// если номера записи не существует, то ничего не происходит, возвращает 0
int CDbf::GoTo(unsigned long recno)
{
	if (!IsInit)
		PutBuffer();
	else
		IsInit = 0;
	if ((recno > 0) && (recno <= lastrec)) {
		_recno = recno;
		ClearBOF_EOF();
		GetBuffer();
		return 1;
	}
	else {
		if (recno == lastrec + 1) {
			SetEOF();
			return 1;
		}
		else
			return 0;
	}
}

unsigned long CDbf::Recno()
{
	return _recno;
}

int CDbf::GetBuffer()
{
	//unsigned long base;
	//base = 32 + 32 * CountField + 2;
	f.seekg(base0 + (_recno - 1) * len_field, ios::beg);
	f.read(_buffer, len_field);
	return 1;
}

int CDbf::ClearBuffer()
{
	int pos = 1;
	memset(_buffer, 32, len_field);
	for (int i = 0; i < this->CountField; i++) {
		if (dbfstruct[i].dbf_type[0] == 'I')
			memset(_buffer + pos, 0, 4);
		pos += dbfstruct[i].len;
	}
	return 1;
}

void CDbf::GoTop()
{
	unsigned long n = 1;
	while (n <= lastrec) {
		GoTo(n);
		if (_buffer[0] != '*') {
			ClearBOF_EOF();
			return;
		}
		n++;
	}
	SetEOF();
}

int CDbf::Deleted()
{
	//MessageInt(_buffer[0]);
	if ((_buffer[0]) == '*')
		return 1;
	else
		return 0;
}

int CDbf::PutBuffer()
{
	//unsigned long base;
	if (_recno > 0) {
		//base = 32 + 32 * CountField + 2;
		f.seekp(base0 + (_recno - 1) * len_field, ios::beg);
		f.write(_buffer, len_field);
		return 1;
	}
	else
		return 0;
}

void CDbf::SetEOF()
{
	_eof = 1;
	_recno = lastrec + 1;
	ClearBuffer();
}

void CDbf::GoBottom()
{
	unsigned long n = lastrec;
	while (n >= 1) {
		GoTo(n);
		if (!Deleted())
			goto ex;
		n--;
	}
ex:
	if (n < 1)
		SetEOF();
	else
		ClearBOF_EOF();
}

void CDbf::Skip(long step) {
	if (step == 0)
		GoTo(_recno);
	if (step > 0)
		Skip1(step);
	if (step < 0)
		Skip2(-step);
}

void CDbf::Skip1(long step)
{
	unsigned long n = _recno + step;
	while (n <= lastrec) {
		GoTo(n);
		if (!Deleted())
			goto ex;
		n++;
	}
ex:
	if (n > lastrec)
		SetEOF();
	else
		ClearBOF_EOF();
}

void CDbf::ClearBOF_EOF()
{
	_eof = 0;
	_bof = 0;
}

void CDbf::Skip2(long step)
{
	unsigned long old_recno = _recno;
	unsigned long n = _recno - step;
	while (n >= 1) {
		GoTo(n);
		if (!Deleted())
			goto ex;
		n--;
	}
ex:
	if (n < 1) {
		GoTo(old_recno);
		_bof = 1;
	}
	else
		ClearBOF_EOF();
}

void CDbf::FieldPut(char *value, char *fieldname)
{
	int pos;
	int idx = GetIndexField(fieldname, pos);
	ostrstream ff(_buffer + pos, dbfstruct[idx].len);
	switch (dbfstruct[idx].dbf_type[0]) {
	case 'C':
		//strncpy(_buffer + pos, value, dbfstruct[idx].len);
		//strncpy(_buffer + pos + strlen(value), value, dbfstruct[idx].len);
		CharToOem(value, value);
		ff << value;
		break;
	case 'N':
		ff << setfill(' ') << setw(dbfstruct[idx].len) << value;
		break;
	case 'L':
		strncpy(_buffer + pos, value, 1);
		break;
	case 'D':
		strncpy(_buffer + pos, value, 8);
		break;
	case 'I':
		strncpy(_buffer + pos, value, 4);
		break;
	}
	PutBuffer();
}

int CDbf::GetIndexField(char *s, int& pos)
{
	int i;
	pos = 1;
	for (i = 0; i < CountField; i++) {
		if (strcmp(dbfstruct[i].dbf_name, s) == 0)
			goto ex;
		pos += dbfstruct[i].len;
	}
ex:
	if (i == CountField)
		return -1;
	else {
		return i;
	}
}

char* CDbf::FieldGet(char *fieldname, char *value, int& lenbuf)
{
	int pos;
	int idx = GetIndexField(fieldname, pos);

	if (idx != -1) {
		if (lenbuf > dbfstruct[idx].len) {
			memcpy(value, _buffer + pos, dbfstruct[idx].len);
			value[dbfstruct[idx].len] = 0;
			OemToChar(value, value);
			return value;
		}
		else {
			lenbuf = dbfstruct[idx].len + 1;
			return NULL;
		}
	}
	else
		return NULL;
}

int CDbf::Eof()
{
	return _eof;
}

char CDbf::FieldType(char *fieldname)
{
	int pos;
	int idx = GetIndexField(fieldname, pos);
	if (idx != -1) {
		return dbfstruct[idx].dbf_type[0];
	}
	else
		return 0;
}

int CDbf::IsOpen()
{
	return f.is_open();
}

void CDbf::Delete()
{
	if (!_eof) {
		_buffer[0] = '*';
		PutBuffer();
	}
}

void CDbf::ReCall()
{
	if (!_eof) {
		_buffer[0] = ' ';
		PutBuffer();
	}
}

int CDbf::Pack()
{
	CDbf oTemp;
	char cTempBuffer[MAX_PATH + 1];
	char cFileName[MAX_PATH + 1];
	strcpy(cFileName, filename);
	GetTempFileName(".", "", 0, cTempBuffer);
	Create(cTempBuffer, dbfstruct, CountField);
	if (oTemp.Open(cTempBuffer)) {
		GoTop();
		while (!_eof) {
			if (_buffer[0] == ' ') {
				oTemp.Append();
				oTemp._FillBuffer(_buffer);
			}
			Skip(1);
		}
		oTemp.Close();
		Close();
		DeleteFile(cFileName);
		MoveFile(cTempBuffer, cFileName);
		Open(cFileName);
		return 1;
	}
	else
		return 0;
}

char* CDbf::NameOfField(int pos)
{
	if ((pos >= 1) && (pos <= CountField))
		return dbfstruct[pos - 1].dbf_name;
	else
		return NULL;
}

int CDbf::FieldCount()
{
	return CountField;
}

void CDbf::_FillBuffer(char *s)
{
	strncpy(_buffer, s, len_field);
	PutBuffer();
}

void CDbf::Zap()
{
	char cTempBuffer[MAX_PATH + 1];
	char fn[MAX_PATH + 1];
	GetTempFileName(".", "", 0, cTempBuffer);
	CDbf::Create(cTempBuffer, dbfstruct, CountField);
	strcpy(fn, filename);
	Close();
	DeleteFile(fn);
	MoveFile(cTempBuffer, fn);
	Open(fn);
}

void CDbf::_putdate1(ofstream& f1)
{
	time_t t;
	struct tm *tm;
	char ch;

	time(&t);
	tm = localtime(&t);

	if (tm->tm_year >= 100)
		ch = (char)tm->tm_year - 100;
	else
		ch = (char)tm->tm_year;
	f1.write((char*)&ch, 1);
	ch = (char)tm->tm_mon + 1;
	f1.write((char*)&ch, 1);
	ch = (char)tm->tm_mday;
	f1.write((char*)&ch, 1);
}

int CDbf::Locate()
{
	unsigned long oldrecno = _recno;
	GoTop();
	while (Eof() != 1) {
		if (LocateCalc() == 1)
			return 1;
		Skip(1);
	}
	GoTo(oldrecno);
	return 0;
}

int CDbf::LocateCalc()
{
	return 0;
}

int CDbf::FieldLen(char* fieldname)
{
	int pos = 0;
	int idx;
	if ((idx = GetIndexField(fieldname, pos)) == -1)
		return -1;
	return dbfstruct[idx].len;
}

void CDbf::FieldGetString(string field, string& value)
{
	int lenbuf = 0;
	char *buf = NULL;
	basic_string<char>::size_type pos;
	FieldGet((char*)field.c_str(), NULL, lenbuf);
	buf = new char[lenbuf];
	FieldGet((char*)field.c_str(), buf, lenbuf);
	value = string(buf);
	delete buf;
	pos = value.find_last_not_of('\x20');
	if (pos == basic_string<char>::npos)
		value = string("");
	else
		value = value.substr(0, pos + 1);
}

int CDbf::FieldGetInt(string field)
{
	int lenbuf = 0;
	char *buf = NULL;
	int rt;
	FieldGet((char*)field.c_str(), NULL, lenbuf);
	buf = new char[lenbuf];
	if (FieldGet((char*)field.c_str(), buf, lenbuf) != NULL)
		rt = atoi(buf);
	else
		rt = 0;
	delete buf;
	return rt;
}


void CDbf::FieldPutString(string field, string value)
{
	FieldPut((char*)value.c_str(), (char*)field.c_str());
}

void CDbf::FieldPutInt(string field, int value)
{
	char s[50];
	_itoa(value, s, 10);
	FieldPut(s, (char*)field.c_str());
}

#pragma once
#include <string>

using namespace std;

int RTrimSpaces(char *s);
string GetEditText(HWND hwnd);
void GetOnlyFilenameFromFullPath(string path, string& fname);
void GetTempFileName(string& s);

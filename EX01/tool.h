//#include"pch.h"
//#include<bits\stdc++.h>
#ifndef TOOL_H
#define TOOL_H

#include<memory>
#include<functional>
#include<string>
#include<vector>
#include<sstream>
//initial tables
//						����,		��ʽ���ַ���			������
using namespace std;
struct pvs :public pair<wstring, pair<wstring, vector<vector<wstring>>>> {
	wstring& tableName = first;
	wstring& formatString = second.first;
	vector<vector<wstring>>& tableData = second.second;
};
struct cls {
	wstring tableName;
	wstring formatString;
	vector<vector<wstring>> tableData;
	size_t colCnt;
};
CString  modifySqlStringPara(const CString& para);
CString dateS2InputDate(CString s);
void generateSQLStringForinsert();

CString getInput(CWnd*pwnd, UINT uid);

#endif
#include"pch.h"
#include"tool.h"
using namespace std;
#pragma error(disable:2362)
CString modifySqlStringPara(const CString& para)
{
	CString ret;
	for (int i = 0; i < para.GetLength(); ++i)
		if (para[i] == '\'')
			ret += "\'\'";
		else
			ret += para[i];
	return ret;
}

CString dateS2InputDate(CString s)
{
	CString ret;
	ret.Format("to_date('%s' , 'YYYY-MM-DD HH24:MI:SS')", s.GetBuffer());
	return ret;
}

void generateSQLStringForinsert() {
	vector<cls> tbs = {
		{ L"BOOKSINFO",L"insert into BOOKSINFO (BOOKNAME,BOOKNO,AUTHOR,PRICE)"\
		L"values('','','','');",{},4u },
		{L"BOOKSIN",L"insert into BOOKSIN (BOOKID,NUM,PKRQ) ('','','');",
		{},3u},
		{L"BOOKSOUT",L"insert into BOOKSOUT (BOOKID,STORAGE)('','');",{},2U}
	};
	vector<wstringstream> ss(tbs.size());
	vector<vector<wstring>> outs(tbs.size());
	ss[0] << LR"(
C#������	10001	���.�׿���˹	199
Linux�ں�Դ�����龰����	10002	ë�²�and��ϣ��	108
Orange'sһ������ϵͳ��ʵ��	10003	��Ԩ	69
������ѧ	10004	�ߵ���.������	99
�������	10005	Jon.Bentley	39
����ǳ��MFC	10006	���	80
STLԴ�����	10007	���	79
��Ϣ�ۻ���	10008	Thomas	58
�˻����������̳�	10009	������and��ѧ��and�����and���	39.5

)";
	ss[1] << LR"(

)";
	ss[2] << LR"()";
	for (int i = 0; i < tbs.size(); ++i) {
		vector<vector<wstring>>& dg = tbs[i].tableData;
		wstringstream& sin = ss[i];
		vector<wstring>& out = outs[i];
		vector<wstring> cache(tbs[i].colCnt);
		while (1) {
			wstring formatString = tbs[i].formatString;
			for (int j = 0; j < cache.size(); ++j)
				if (!(sin >> cache[j]))
				{
					if(j!=0)
				err:
					AfxMessageBox(CString("you see ,sin is not fit the cols as expected,")
						+ "while generate SQLStirng for table "
						+ CString(tbs[i].tableName.c_str()));
					goto nexttb;
				}
			int pos = 0;
			for (auto& para : cache)
			{
				{
					wstring p;
					for (auto c : para)
						if (c == L'\'')
							p += L"\'\'";
						else 
							p += c;
					para = p;
				}

				pos = formatString.find(L"''", pos);
				if (pos == static_cast<size_t>(-1))goto err;
				formatString.insert(pos + 1, para);
				pos += 2 + para.length();
			}
			out.push_back(formatString);
			cache.resize(4);
		}
	nexttb:;
	}
	//�����
	wstringstream ssout;
	ssout << "SET define off";
	for (auto out : outs) {
		ssout << "\n\n";
		for (auto o : out)
			ssout << o << "\n";
	}
	//notice me
	wstring wsout = ssout.str();
	return;
}

CString getInput(CWnd* pwnd, UINT uid)throw(CString)
{
	auto p = pwnd->GetDlgItem(uid);
	if (!p) {
		throw CString("��Ч�ؼ�");
	}
	CString s;
	p->GetWindowTextA(s);
	s = s.Trim();
	if (s.IsEmpty()) {
		p->SetFocus();
	}
	return s;
}

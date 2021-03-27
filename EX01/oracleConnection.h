#pragma once
#include"mOracle.h"
#include<string>
#include<afxwin.h>
//һ�����������һ�����ݿ�����
class oracleConnection
{
public :
	using recSetType = _RecordsetPtr;//��¼����������
	using connectType = _RecordsetPtr;//���Ӷ�������
	using stringType = CString;//�ַ�������
	enum class returnCode { right = 0, error = 1 };
	//using  enum returnCode; //not work until c++20 
	connectType pDBConnection; // ���ݿ� ����

	//��ʼ��
	oracleConnection();
	oracleConnection(stringType connectSql);
	/// <summary>
	/// �������ݿ����ӣ�Ҳ�ɳ�ʼ��
	/// </summary>
	/// <param name="connectSql">�����ַ���</param>
	/// <returns></returns>
	int openConnection(stringType connectSql);
};


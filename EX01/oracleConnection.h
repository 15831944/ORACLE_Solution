#pragma once
#include"mOracle.h"
#include<string>
#include<afxwin.h>
//一个对象代表着一个数据库连接
class oracleConnection
{
public :
	using recSetType = _RecordsetPtr;//记录集对象类型
	using connectType = _RecordsetPtr;//连接对象类型
	using stringType = CString;//字符串类型
	enum class returnCode { right = 0, error = 1 };
	//using  enum returnCode; //not work until c++20 
	connectType pDBConnection; // 数据库 连接

	//初始化
	oracleConnection();
	oracleConnection(stringType connectSql);
	/// <summary>
	/// 更改数据库连接，也可初始化
	/// </summary>
	/// <param name="connectSql">连接字符串</param>
	/// <returns></returns>
	int openConnection(stringType connectSql);
};


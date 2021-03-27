#pragma once

#ifndef    __DB_ORACLE
#define    __DB_ORACLE

#include "pch.h"
#include "afxwin.h"

class TZYDB
{
public:
	  static _ConnectionPtr      pDBConnection; // 数据库 连接
      static _RecordsetPtr       pDBRecordSet; // 记录
      static _CommandPtr         pDBCommand; // 命令 	

	  TZYDB(void);
	  ~TZYDB(void);
	  
public:	
	  //static bool OnInitADOConn();
	  static bool OnInitADOConn(CString linkStr="");
      static void closeDB();
	  static void closeConnetion();
	 
	  //执行的是 select 指令
	  static _RecordsetPtr openDB(CString sSql, bool bShowErr = true);

	  static _RecordsetPtr openForQuikRead(CString sSql, bool bShowErr = true);

	  //获取数据集的字段名
	  static void getFiledName(_RecordsetPtr pDataSet, CStringArray &aFiledName);

	  //基于Sql查询，得到的数据填写到列表控件pListCtrl中
	  static bool dataToListCtrl(CListCtrl *pListCtrl, CString sSql);

	  //第三周新增 start	  
	  //数据库增、删、改 insert、delete、update、create、drop
	  static int executeSql(CString sSql);

	   //基于事务的数据库增、删、改 insert、delete、update、create、drop
	  static int  execTransaction(const CStringArray &aSql);

	  //保存二进制（ole）数据,数据来源于指针
	  static bool saveOleData(CString sSql, byte *srcData, long srcDataBytes, CString oleFiledName);

	  //保存二进制（ole）数据,数据来源于文件
	  static bool saveOleFromFile(CString sSql, CString sFilePath, CString oleFiledName);

	  //获取ole数据的字节数
	  static long getOleBytes(CString sSql, CString oleFiledName);

	  //从打开的数据集中取出ole数据
	  static byte *oleToBufferByOpenDB(CString oleFiledName, long &dataBytes);

	  //读ole数据，成功返回数据的首地址，否则返回NULL,注意内存由主调函数释放
	  static byte *readOleData(CString sSql, CString oleFiledName, long &dataBytes);
	  //第三周新增End

	  //第五周

	  //select count(*)...
	   static int getRecordCount(CString sSql);

	   //select password...
	   static CString getOneItem(CString sSql);
	  //第五周End


	   //第六周 放在 Oracle.h中
	   static void dataSetToListCtrl(CListCtrl *pListCtrl,_RecordsetPtr pDataSet);

	   /*基于存储过程获取数据集,sCommandText应给出存储过程名及参数，但参数不包含第一个输出游标
	    如：create or replace procedure ComHisStorage_2(sys_cur out sys_refcursor, byDate in date) 
		则：sCommandText应为：ComHisStorage_2(to_date('2020-03-03','yyyy-mm-dd'))
	   */
	   static _RecordsetPtr getpRecSetByStoredProc(CString sCommandText);
	   //第六周End

	   //第八周 放在 Oracle.h中
	   //强制断开用户 sUserName 的连接
	   static bool disconnectUser(CString sUserName, bool m_bShowMes = true);
	   //第八周End


public://一些控件处理函数
	//清空表格控件
	static void clearCListCtrl(CListCtrl *pListCtrl);

	//填写表格控件的头
	static void doListHeader(CListCtrl *pListCtrl, const CStringArray &aColName);

	//自动设置表格控件列宽
	static void listAutoWidth(CListCtrl *pListCtrl, int nColSpace = 0, int nMinWid = 0);

	 //第三周新增图像功能 start	  
	static void clearPic(CStatic *pPic, COLORREF *pColor = NULL);
	static bool loadImgFromFile(CDC *dstDC, CRect &dstRec,  CString sFile);
	static bool showImgByImagePtr(CDC *dstDC, CRect dstRec, CImage *pImage, bool bRevise = true);	
	static CImage *dbImgToCImage(byte *pMemData, long dataBytes);
	static CImage *dbImgToCImage(CString sSql, CString oleFiledName);
	static bool showDBImg(CDC *dstDC, CRect dstRec,  CString sSql, CString oleFiledName);
	static void showDBImage(CString sSql, CString oleName, CStatic *pic);
	 //第三周新增图像功能 end 
};

#endif
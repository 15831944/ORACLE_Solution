#pragma once

#ifndef    __DB_ORACLE
#define    __DB_ORACLE

#include "pch.h"
#include "afxwin.h"

class TZYDB
{
public:
	  static _ConnectionPtr      pDBConnection; // ���ݿ� ����
      static _RecordsetPtr       pDBRecordSet; // ��¼
      static _CommandPtr         pDBCommand; // ���� 	

	  TZYDB(void);
	  ~TZYDB(void);
	  
public:	
	  //static bool OnInitADOConn();
	  static bool OnInitADOConn(CString linkStr="");
      static void closeDB();
	  static void closeConnetion();
	 
	  //ִ�е��� select ָ��
	  static _RecordsetPtr openDB(CString sSql, bool bShowErr = true);

	  static _RecordsetPtr openForQuikRead(CString sSql, bool bShowErr = true);

	  //��ȡ���ݼ����ֶ���
	  static void getFiledName(_RecordsetPtr pDataSet, CStringArray &aFiledName);

	  //����Sql��ѯ���õ���������д���б�ؼ�pListCtrl��
	  static bool dataToListCtrl(CListCtrl *pListCtrl, CString sSql);

	  //���������� start	  
	  //���ݿ�����ɾ���� insert��delete��update��create��drop
	  static int executeSql(CString sSql);

	   //������������ݿ�����ɾ���� insert��delete��update��create��drop
	  static int  execTransaction(const CStringArray &aSql);

	  //��������ƣ�ole������,������Դ��ָ��
	  static bool saveOleData(CString sSql, byte *srcData, long srcDataBytes, CString oleFiledName);

	  //��������ƣ�ole������,������Դ���ļ�
	  static bool saveOleFromFile(CString sSql, CString sFilePath, CString oleFiledName);

	  //��ȡole���ݵ��ֽ���
	  static long getOleBytes(CString sSql, CString oleFiledName);

	  //�Ӵ򿪵����ݼ���ȡ��ole����
	  static byte *oleToBufferByOpenDB(CString oleFiledName, long &dataBytes);

	  //��ole���ݣ��ɹ��������ݵ��׵�ַ�����򷵻�NULL,ע���ڴ������������ͷ�
	  static byte *readOleData(CString sSql, CString oleFiledName, long &dataBytes);
	  //����������End

	  //������

	  //select count(*)...
	   static int getRecordCount(CString sSql);

	   //select password...
	   static CString getOneItem(CString sSql);
	  //������End


	   //������ ���� Oracle.h��
	   static void dataSetToListCtrl(CListCtrl *pListCtrl,_RecordsetPtr pDataSet);

	   /*���ڴ洢���̻�ȡ���ݼ�,sCommandTextӦ�����洢����������������������������һ������α�
	    �磺create or replace procedure ComHisStorage_2(sys_cur out sys_refcursor, byDate in date) 
		��sCommandTextӦΪ��ComHisStorage_2(to_date('2020-03-03','yyyy-mm-dd'))
	   */
	   static _RecordsetPtr getpRecSetByStoredProc(CString sCommandText);
	   //������End

	   //�ڰ��� ���� Oracle.h��
	   //ǿ�ƶϿ��û� sUserName ������
	   static bool disconnectUser(CString sUserName, bool m_bShowMes = true);
	   //�ڰ���End


public://һЩ�ؼ�������
	//��ձ��ؼ�
	static void clearCListCtrl(CListCtrl *pListCtrl);

	//��д���ؼ���ͷ
	static void doListHeader(CListCtrl *pListCtrl, const CStringArray &aColName);

	//�Զ����ñ��ؼ��п�
	static void listAutoWidth(CListCtrl *pListCtrl, int nColSpace = 0, int nMinWid = 0);

	 //����������ͼ���� start	  
	static void clearPic(CStatic *pPic, COLORREF *pColor = NULL);
	static bool loadImgFromFile(CDC *dstDC, CRect &dstRec,  CString sFile);
	static bool showImgByImagePtr(CDC *dstDC, CRect dstRec, CImage *pImage, bool bRevise = true);	
	static CImage *dbImgToCImage(byte *pMemData, long dataBytes);
	static CImage *dbImgToCImage(CString sSql, CString oleFiledName);
	static bool showDBImg(CDC *dstDC, CRect dstRec,  CString sSql, CString oleFiledName);
	static void showDBImage(CString sSql, CString oleName, CStatic *pic);
	 //����������ͼ���� end 
};

#endif
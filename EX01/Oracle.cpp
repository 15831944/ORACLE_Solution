#include "pch.h"
#include "Oracle.h"
#include "MyControls.h"
//------------------------------------------------------------------------
_ConnectionPtr    TZYDB::pDBConnection = NULL; // 数据库 
_RecordsetPtr     TZYDB::pDBRecordSet  = NULL; // 命令 
_CommandPtr       TZYDB::pDBCommand = NULL; // 记录
//------------------------------------------------------------------------
//第三周新增图像功能 start	  
bool TZYDB::loadImgFromFile(CDC *dstDC, CRect &dstRec,  CString sFile){
	try{
		CImage img;
		if(!img.IsNull()){
			img.Destroy();
		}

		img.Load(sFile);	   
		
		if (img.IsNull()){
			throw CString("加载图像：" + sFile + "  失败");
		}

		int imgW = img.GetWidth();
		int imgH = img.GetHeight();
		
		CDC memDC;
		memDC.CreateCompatibleDC(NULL);
		CBitmap memBmp;
		memBmp.CreateCompatibleBitmap(dstDC,  imgW,  imgH);
		memDC.SelectObject(&memBmp);
		img.Draw(memDC.m_hDC, 0, 0);
		img.Destroy();
		
		if(dstRec.Width() < imgW || dstRec.Height() < imgH) {
			double zoom = (double)dstRec.Width() / (double)imgW;
			double zoomH = (double)dstRec.Height() / (double)imgH;
			if(zoom > zoomH) zoom = zoomH;
			int dstW = imgW * zoom, dstH = imgH * zoom;
			dstDC->StretchBlt(0, 0, dstW, dstH, &memDC, 0, 0, imgW, imgH, SRCCOPY);
		} else {
			dstDC->BitBlt(0, 0, dstRec.Width(),  dstRec.Height(), &memDC, 0, 0, SRCCOPY);
		}	

		memBmp.DeleteObject();
		memDC.DeleteDC();

		return true;	
	} catch(CString &eStr) {
		AfxMessageBox(eStr);	
	} catch(...) {
		AfxMessageBox("TZYDB::loadImgFromFile Error");	
	}

	return false;
}
//--------------------------------------------------------------------------
bool TZYDB::showImgByImagePtr(CDC *dstDC, CRect dstRec, CImage *pImage, bool bRevise){
	if(pImage == NULL) {
		AfxMessageBox("TZYDB::showImgByImagePtr\r\n pImage == NULL");
		return false;
	}

	if(bRevise) {
		dstRec.top += 1;
		dstRec.left += 1;
		dstRec.right -= 1;
		dstRec.bottom -= 1;
	}

	int imgW = pImage->GetWidth();
	int imgH = pImage->GetHeight();
	
	CDC memDC;
	memDC.CreateCompatibleDC(NULL);
	CBitmap memBmp;
	
	if(imgW < dstRec.Width()){
		imgW = dstRec.Width();
	}

	if(imgH < dstRec.Height()){
		imgH = dstRec.Height();
	}

	memBmp.CreateCompatibleBitmap(dstDC,  imgW,  imgH);
	memDC.SelectObject(&memBmp);
	memDC.FillSolidRect(&dstRec,RGB(240,240,240)); 
	memDC.Rectangle(0, 0, dstRec.Width(), dstRec.Height());
	pImage->Draw(memDC.m_hDC, 0, 0);

	if(bRevise == false) {	
		CRect srcRec(0,0,pImage->GetWidth(),pImage->GetHeight());
		pImage->StretchBlt(dstDC->m_hDC, dstRec, srcRec, SRCCOPY);
	} else if(dstRec.Width()-2 <= imgW || dstRec.Height()-2 <= imgH) {
		double zoom = (double)dstRec.Width() / (double)imgW;
		double zoomH = (double)dstRec.Height() / (double)imgH;

		if(zoom > zoomH){
			zoom = zoomH;
		}

		int dstW = imgW * zoom - 2;
		int dstH = imgH * zoom - 2;

		int left = (dstRec.Width() - dstW) / 2 + 1;
		int top = (dstRec.Height() - dstH) / 2 + 1;
		dstDC->StretchBlt(left, top, dstW, dstH, &memDC, 0, 0, imgW, imgH, SRCCOPY);
	} else {
		int left = (dstRec.Width() - imgW) / 2 + 1;
		int top = (dstRec.Height() - imgH) / 2;

		dstDC->BitBlt(left, top, dstRec.Width(),  dstRec.Height(), &memDC, 0, 0, SRCCOPY);
	}		 
	
	memBmp.DeleteObject();
	memDC.DeleteDC();
	
	return true;

}
//--------------------------------------------------------------------------
CImage *TZYDB::dbImgToCImage(byte *pMemData, long dataBytes) {
	CString sError = "CImage *ZYFunction::dbImgToCImage\r\n", sFlag = "";
	IStream *pStream  =  NULL;
	HGLOBAL hGlobal = NULL;

	hGlobal = GlobalAlloc(GMEM_MOVEABLE,  dataBytes);
	if(hGlobal == NULL){
		CString s;
		s.Format("GlobalAlloc(GMEM_MOVEABLE, %d) 失败", dataBytes);
		throw CString(sError + s);	
	}

	void *pData = GlobalLock(hGlobal);
	memcpy(pData, pMemData, dataBytes);
	GlobalUnlock(hGlobal);

	if(CreateStreamOnHGlobal(hGlobal, TRUE,  &pStream) != S_OK){
		throw CString(sError + "CreateStreamOnHGlobal失败");
	}

	CImage *pImg = new CImage();
	if(pImg == NULL){
		sFlag = "pImg == NULL\r\nnew CImage()失败";
	}else if (SUCCEEDED(pImg->Load(pStream))){
		if (pImg->IsNull()){
			sFlag = "加载流后图像为空";
		}
	} else {
		sFlag = "pImg->Load(pStream)失败";
	}

	if(pStream){
		pStream->Release(); 
	}

	if(hGlobal){
		GlobalFree(hGlobal);
	}

	if(sFlag.GetLength() > 3){
		throw CString(sError + sFlag);
	}

	return pImg;
}
//--------------------------------------------------------------------------
CImage *TZYDB::dbImgToCImage(CString sSql, CString oleFiledName) {
	long dataBytes;
	byte *pMemData = TZYDB::readOleData(sSql, oleFiledName, dataBytes);

	if(!pMemData){
		throw CString("CImage *ZYFunction::dbImgToCImage\r\n读取Ole数据失败\n" + sSql);
	}

	CImage *pImage = dbImgToCImage(pMemData, dataBytes);
	delete[] pMemData;
	return pImage;
}
//--------------------------------------------------------------------------
bool TZYDB::showDBImg(CDC *dstDC, CRect dstRec,  CString sSql, CString oleFiledName){
	try{
		CImage *pImage = dbImgToCImage(sSql, oleFiledName);
		showImgByImagePtr(dstDC, dstRec, pImage);
		pImage->Destroy();

		delete pImage;
		return true;
	}catch(CString &eStr){
		AfxMessageBox("TZYDB::showDBImg\n" + eStr);
	}catch(...){
		AfxMessageBox("TZYDB::showDBImg:  未知错误");	         
	}
	return false;

}
//--------------------------------------------------------------------------
void TZYDB::clearPic(CStatic *pPic, COLORREF *pColor){
	CRect rec;
	COLORREF bkColor = (pColor ? RGB(249,249,244) : RGB(240,240,240));

	CDC *dc = pPic->GetDC();
	pPic->GetClientRect(&rec);
	
	dc->Rectangle(rec.left, rec.top, rec.right, rec.bottom);

	pPic->ReleaseDC(dc);	
}
//--------------------------------------------------------------------------
void TZYDB::showDBImage(CString sSql, CString oleName, CStatic *pic){
	COLORREF color=RGB(255,255,255);

	clearPic(pic, &color);
	
	CRect rec;
	
	pic->GetClientRect(&rec);
	rec.left += 1;
	CDC *pDC = pic->GetDC();
	showDBImg(pDC, rec,  sSql,  oleName);

	pic->ReleaseDC(pDC);
}
	 //第三周新增图像功能 end 
//--------------------------------------------------------------------------
//第5周新增功能 
//--------------------------------------------------------------------------
int TZYDB::getRecordCount(CString sSql) {
	try {
		if(openDB(sSql) == NULL) return -1;
		_variant_t vIndex = (long)0;
		_variant_t vCount = pDBRecordSet->GetCollect(vIndex);
		closeDB();
		return vCount.lVal;
	}catch(_com_error e){		
		AfxMessageBox("TZYDB::getRecordCount 失败\n" + CString((LPCTSTR)e.Source()) + TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));//+ e.ErrorInfo());
		closeDB();
		return false;
	}
}
//--------------------------------------------------------------------------
CString TZYDB::getOneItem(CString sSql) {
	try {
		if(openDB(sSql) == NULL || pDBRecordSet->adoEOF){
			closeDB();
			return "";
		}
		_variant_t vIndex = (long)0;
		_variant_t var = pDBRecordSet->GetCollect(vIndex);
		CString s = "";
		if(var.vt != VT_NULL) s = (LPCSTR)_bstr_t(var);
		closeDB();
		return s;
	} catch(_com_error e){		
		AfxMessageBox("TZYDB::getOneItem 失败\n" + CString((LPCTSTR)e.Source()) + TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));//+ e.ErrorInfo());
		closeDB();
	}
	return "";
}
//--------------------------------------------------------------------------
//第5周新增功能 End
//--------------------------------------------------------------------------

//--------------------------------------------------------------------------
//第六周 Start
void TZYDB::dataSetToListCtrl(CListCtrl *pListCtrl, _RecordsetPtr pRecSet){
	try {		
		pListCtrl->SetExtendedStyle(pListCtrl->GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES );   
		
	
		CStringArray aFiledName;

		//获取数据集的字段名
	    getFiledName(pRecSet, aFiledName);
		
		CString s = "";
		int iCount = aFiledName.GetSize();
	

		//空间暂时不要绘制，否则数据量大的时候会很慢
		pListCtrl->SetRedraw(FALSE);
		
		//清空表格控件
		clearCListCtrl(pListCtrl);

		//填写表格控件的头
		doListHeader(pListCtrl, aFiledName);

		//遍历数据集，制表
		_variant_t var;
		int rowIndex = 0;
		int nCol = 0;

		while(!pRecSet->adoEOF) {
			nCol = 0;
			for( int i = 0; i < iCount; i++){				
				var = pRecSet->GetCollect(aFiledName[i].GetBuffer());
				s = (var.vt == VT_NULL ? "" : (LPCSTR)_bstr_t(var));				

				nCol == 0 ? pListCtrl->InsertItem(rowIndex,  s) : pListCtrl->SetItemText(rowIndex,  nCol,  s);
				nCol++;
			}

			rowIndex++;
			
			pRecSet->MoveNext();
		}

		pRecSet->Close();
		listAutoWidth(pListCtrl, 13, 80);

		pListCtrl->SetRedraw(TRUE);
		pListCtrl->InvalidateRect(NULL, FALSE);
	} catch(CString &eStr) {
		AfxMessageBox("TZYDB::dataToListTable 失败\n" + eStr);
	}catch(_com_error e){		
		AfxMessageBox("TZYDB::dataToListTable 失败\n" + CString((LPCTSTR)e.Source()) + TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));//+ e.ErrorInfo());
		closeDB();
	}

	pListCtrl->SetRedraw(TRUE);
}
//--------------------------------------------------------------------------
/*基于存储过程获取数据集,sCommandText应给出存储过程名及参数，但参数不包含第一个输出游标
	    如：create or replace procedure ComHisStorage_2(sys_cur out sys_refcursor, byDate in date) 
		则：sCommandText应为：ComHisStorage_2(to_date('2020-03-03','yyyy-mm-dd'))
*/
_RecordsetPtr TZYDB::getpRecSetByStoredProc(CString sCommandText){
	try{
		if(!TZYDB::OnInitADOConn()){
			return NULL;  
		}

		_CommandPtr m_pCommand;

		m_pCommand.CreateInstance(__uuidof(Command));		
		
		m_pCommand->ActiveConnection = TZYDB::pDBConnection;

		//存储过程及参数
		m_pCommand->CommandText = sCommandText.AllocSysString(); 
		
		m_pCommand->CommandType = adCmdStoredProc;
		

		if (pDBRecordSet && pDBRecordSet->GetState() == adStateOpen){
			pDBRecordSet->Close();
		}
	
        TZYDB::pDBRecordSet = m_pCommand->Execute(NULL,NULL,adCmdStoredProc | adCmdUnspecified);	
     
		//返回的数据集应该由主调函数关闭，以释放内存
		return TZYDB::pDBRecordSet;	
	}catch (_com_error &e){
			AfxMessageBox("" + CString((LPCTSTR)e.Source()) +
				TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));	  
	}

	return NULL;
}
//第六周End
//--------------------------------------------------------------------------

///////////////////////控件处理函数（与数据库无直接关系）
void TZYDB::clearCListCtrl(CListCtrl *pListCtrl) {
	int nColumnCount = 0;
	//加个判断
	if(pListCtrl->GetHeaderCtrl()){
		nColumnCount = pListCtrl->GetHeaderCtrl()->GetItemCount();
	}

	pListCtrl->DeleteAllItems(); 

	for (int i = 0;i < nColumnCount; i++) {
		pListCtrl->DeleteColumn(0);
	}
}
//--------------------------------------------------------------------------
void TZYDB::doListHeader(CListCtrl *pListCtrl, const CStringArray &aColName){
	if(pListCtrl->GetHeaderCtrl()){
		clearCListCtrl(pListCtrl);
	}

	//InsertColumn 不是虚函数，需要手工动态进行指针类型转换
	CMyCListCtrl *m_pMyList = dynamic_cast<CMyCListCtrl *>(pListCtrl);

	pListCtrl->SetExtendedStyle(pListCtrl->GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER); 
	
	for(int i = 0; i < aColName.GetSize(); i++) {
		if(m_pMyList){
			m_pMyList->InsertColumn(i + 1, (LPCTSTR)aColName[i], LVCFMT_CENTER, 80); 		
		}else{
			pListCtrl->InsertColumn(i + 1, (LPCTSTR)aColName[i], LVCFMT_CENTER, 80); 		
		}
	}
}
//--------------------------------------------------------------------------
void TZYDB::listAutoWidth(CListCtrl *pListCtrl, int nColSpace, int nMinWid){
	pListCtrl->SetRedraw(FALSE);
	
	int nColumnWidth, nHeaderWidth;
	int nColumnCount = 0;
	
	//获取表格列数
	CHeaderCtrl* pHeaderCtrl = pListCtrl->GetHeaderCtrl();
	if(!pHeaderCtrl){//无表头
		return;
	}
	nColumnCount = pHeaderCtrl->GetItemCount();
	if(nColumnCount < 1){//无列表
		return;
	}
	
	for (int i = 0; i < nColumnCount - 1; i++){	
		pListCtrl->SetColumnWidth(i, LVSCW_AUTOSIZE); 
		
		nColumnWidth = pListCtrl->GetColumnWidth(i) + nColSpace;
		pListCtrl->SetColumnWidth(i, LVSCW_AUTOSIZE_USEHEADER);
		nHeaderWidth = pListCtrl->GetColumnWidth(i) + nColSpace; 
		
		if(nColSpace > 0){
			nColumnWidth = max(nColumnWidth, nHeaderWidth);
		}else{
			nColumnWidth = max(nColumnWidth, nHeaderWidth);
		}

		if(nColumnWidth < nMinWid){
			nColumnWidth = nMinWid;
		}

		pListCtrl->SetColumnWidth(i, nColumnWidth);		
	}

	pListCtrl->SetRedraw(TRUE);

	pListCtrl->InvalidateRect(NULL, FALSE);
}

//------------------------------------------------------------------------
//------------------------------------------------------------------------
///////////////////////以下是与数据库直接相关的函数
bool TZYDB::OnInitADOConn(CString linkStr) {//程序开始建立连接
	static bool isFirst = true;//好东西

	if(linkStr.IsEmpty()){
		linkStr = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(CID=GTU_APP)"
			"(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=127.0.0.1)(PORT=1521)))"
			"(CONNECT_DATA=(SID=orcl)(SERVER=DEDICATED)));"
			"User Id=student;Password=123456;PLSQLRSet=1;";
	}

	if(!isFirst){//只需要连接一次
		return true;
	}

	isFirst = false;	
	
	try {		
		if(FAILED(::CoInitialize(NULL))){
			throw CString("类工厂初始化失败！");
	    }
	
		pDBConnection.CreateInstance("ADODB.Connection");

		_bstr_t bstrConnect = _T(linkStr.GetBuffer());

		//2020-04-22 新增 
		//使用 客户端游标，否则x64的程序 遍历数据集可能只能访问到第一条记录
		//设置游标方式 必须 在 pDBConnection->Open 之前执行
		//另外个 不带参的 OnInitADOConn 也做同样处理
		pDBConnection->CursorLocation = adUseClient;
		//2020-04-22 新增 End

		//注意有的机子，若不是用管理员身份运行，此处可能无法建立连接
		pDBConnection->Open(bstrConnect, "","",adModeUnknown);

		pDBRecordSet.CreateInstance("ADODB.Recordset");
		
		pDBCommand.CreateInstance(__uuidof(Command));
		
		pDBCommand->ActiveConnection = pDBConnection;
		
		return true;
	}catch(_com_error e){
		AfxMessageBox("AdoAccess::OnInitADOConn 连接失败\n" + CString((LPCTSTR)e.Source()) +
			TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));	 		
	}catch(CString &eStr){
		AfxMessageBox("AdoAccess::OnInitADOConn 连接失败\n" + eStr);
	}
	
	return false;
}

//------------------------------------------------------------------------
void TZYDB::closeConnetion(){//程序结束，调用，关闭连接
	try{
		if(pDBRecordSet){
			if(pDBRecordSet->GetState() == adStateOpen){
				pDBRecordSet->Close();
			}

			pDBRecordSet.Release();
			pDBRecordSet = NULL;
		}

		if(pDBCommand){
			pDBCommand.Release();
			pDBCommand = NULL;
		}

		if (pDBConnection){
			if(TZYDB::pDBConnection->GetState() == adStateOpen){
				pDBConnection->Close();
			}

			pDBConnection.Release();
			pDBConnection = NULL;
		}

	
		::CoUninitialize();
	}catch(_com_error e){
		AfxMessageBox("TZYDB::closeConnetion\r\n" + CString((LPCTSTR)e.Source()) +
			TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));	  
	}catch(...){
	    AfxMessageBox("TZYDB::closeConnetion\r\n未知错误");
	}
}
//------------------------------------------------------------------------
void TZYDB::closeDB() {	
	if (pDBRecordSet && pDBRecordSet->GetState() == adStateOpen){
		try{
			pDBRecordSet->Close();
		}catch(_com_error e){
			AfxMessageBox("TZYDB::closeDB\r\n" + CString((LPCTSTR)e.Source()) +
				TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));	  
		}catch(...){
			AfxMessageBox("TZYDB::closeDB()\r\n未知错误");
		}
	}	
}
//------------------------------------------------------------------------
_RecordsetPtr TZYDB::openDB(CString sSql, bool bShowErr) {
	try {
		if(!OnInitADOConn()){
			return NULL;
		}
		//2020-4-28新增
		if(pDBRecordSet->GetState() == adStateOpen){
			pDBRecordSet->Close();
		}
		//2020-4-28新增 End
		pDBRecordSet->Open(
			sSql.GetBuffer(),
			pDBConnection.GetInterfacePtr(),
			adOpenDynamic, 
			adLockOptimistic, 
			adCmdText
			);

		return pDBRecordSet;
	}catch(_com_error e){	
		if(bShowErr){
			AfxMessageBox("TZYDB::openDB 失败\n" + CString((LPCTSTR)e.Source()) + TEXT("\nDescription:\n ") + 
				CString((LPCTSTR)e.Description()) + "\n" + sSql);
		}

		closeDB();
		
		return NULL;
	}
}
//------------------------------------------------------------------------
_RecordsetPtr TZYDB::openForQuikRead(CString sSql, bool bShowErr) {
	try {
		if(!OnInitADOConn()){
			return NULL;
		}
		//2020-4-28新增
		if(pDBRecordSet->GetState() == adStateOpen){
			pDBRecordSet->Close();
		}
		//2020-4-28新增 End
		pDBRecordSet->Open(
			sSql.GetBuffer(),
			pDBConnection.GetInterfacePtr(),
			adOpenForwardOnly, 
			adLockReadOnly, 
			adCmdText
			);

		return pDBRecordSet;
	}catch(_com_error e){	
		if(bShowErr){
			AfxMessageBox("TZYDB::openDB 失败\n" + CString((LPCTSTR)e.Source()) + TEXT("\nDescription:\n ") + 
				CString((LPCTSTR)e.Description()) + "\n" + sSql);
		}

		closeDB();
		
		return NULL;
	}
}


//------------------------------------------------------------------------
void TZYDB::getFiledName(_RecordsetPtr pDataSet, CStringArray &aFiledName) {
	aFiledName.RemoveAll();
	
	int iCount = pDataSet->Fields->Count;
	
	FieldsPtr fields = pDataSet->GetFields(); 

	for(long i = 0; i < iCount; i++){
		aFiledName.Add((LPCSTR)(fields->GetItem(i)->Name));
	}
}
//------------------------------------------------------------------------
bool TZYDB::dataToListCtrl(CListCtrl *pListCtrl, CString sSql){
	try {		
		pListCtrl->SetExtendedStyle(pListCtrl->GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES );   
		
		if(openDB(sSql, true) == NULL){
			return false;
		}

		CStringArray aFiledName;

		//获取数据集的字段名
	    getFiledName(pDBRecordSet, aFiledName);
		
		CString s = "";
		int iCount = aFiledName.GetSize();
	

		//空间暂时不要绘制，否则数据量大的时候会很慢
		pListCtrl->SetRedraw(FALSE);
		
		//清空表格控件
		clearCListCtrl(pListCtrl);

		//填写表格控件的头
		doListHeader(pListCtrl, aFiledName);

		//遍历数据集，制表
		_variant_t var;
		int rowIndex = 0;
		int nCol = 0;

		while(!pDBRecordSet->adoEOF) {
			nCol = 0;
			for( int i = 0; i < iCount; i++){				
				var = pDBRecordSet->GetCollect(aFiledName[i].GetBuffer());
				s = (var.vt == VT_NULL ? "" : (LPCSTR)_bstr_t(var));				

				nCol == 0 ? pListCtrl->InsertItem(rowIndex,  s) : pListCtrl->SetItemText(rowIndex,  nCol,  s);
				nCol++;
			}

			rowIndex++;
			
			pDBRecordSet->MoveNext();
		}

		closeDB();
		listAutoWidth(pListCtrl, 13, 80);

		pListCtrl->SetRedraw(TRUE);
		pListCtrl->InvalidateRect(NULL, FALSE);

		return true;
	} catch(CString &eStr) {
		AfxMessageBox("TZYDB::dataToListTable 失败\n" + eStr);
	}catch(_com_error e){		
		AfxMessageBox("TZYDB::dataToListTable 失败\n" + CString((LPCTSTR)e.Source()) + TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));//+ e.ErrorInfo());
		closeDB();
	}

	pListCtrl->SetRedraw(TRUE);
	
	return false;
}
//-------------------------------------------------------------
 //第三周新增 start	  
 // 插入成功返回受影响行数，否则是-1
//------------------------------------------------------------------------
int TZYDB::executeSql(CString sSql) {
	try {
		if(!OnInitADOConn()){
			return -1;  
		}

		_variant_t recordsAffected[800];
		
		//执行写操作 注:若是读：pDBRecordSet->Open
		pDBConnection->Execute(sSql.GetBuffer(), recordsAffected, 0);
		char buffer[200];
		
		sprintf_s(buffer, "%s", (char*)((_bstr_t)recordsAffected));
		
		int iRecs = atoi(buffer);
		
		//closeConnetion();

		return iRecs;
	}catch(_com_error e){		
		AfxMessageBox("TZYDB::executeSql 失败\n" + CString((LPCTSTR)e.Source()) + TEXT("\n描述:\n ") + CString((LPCTSTR)e.Description()));//+ e.ErrorInfo());
		AfxMessageBox(sSql);
		//closeConnetion();
		return -1;
	}
}
//------------------------------------------------------------------------
int TZYDB::execTransaction(const CStringArray &aSql){
	CString s;

	try {
		if(!OnInitADOConn()){
			return 0;
		}

		_variant_t recordsAffected[200];
		char buffer[200];	
		int nRecs = 0;

		pDBCommand->ActiveConnection->BeginTrans();

		for(int i = 0; i < aSql.GetSize(); i++){
			pDBCommand->CommandText = aSql[i].AllocSysString();

			pDBCommand->Execute(recordsAffected, NULL, adCmdText);	

			sprintf_s(buffer, "%s", (char*)((_bstr_t)recordsAffected));

			s = aSql[i];
			s.MakeUpper();
			nRecs += atoi(buffer);
		}

		pDBCommand->ActiveConnection->CommitTrans();

		return nRecs;
	}catch(_com_error e){
		pDBCommand->ActiveConnection->RollbackTrans();
		s = "执行事务失败，事务已回退。\r\n";
		s += CString((LPCTSTR)e.Source()) + "\r\n错误描述:\r\n" + CString((LPCTSTR)e.Description());

		//MessageBoxA(NULL,"事务操作失败...","", MB_OK);
		MessageBoxA(NULL,s,"事务操作失败...",MB_OK);
		//closeConnetion();

		return 0;
	}
}
//------------------------------------------------------------------------
//调用示例：TZYDB::saveOleData("select photo from BooksInfor where ID=XXX", srcData, 字节数, "photo")
bool TZYDB::saveOleData(CString sSql, byte *srcData, long srcDataBytes, CString oleFiledName) {
	try  {
		if(!OnInitADOConn()){
			return false;
		}

		if(!srcData){
			return false;
		}

		_variant_t  _vl;
		_vl = (_variant_t)(LPCTSTR)sSql;
		pDBRecordSet->Open(_vl, _variant_t((IDispatch*)pDBConnection, true), 
			adOpenStatic, adLockOptimistic, adCmdText);

		VARIANT           varBLOB;
		SAFEARRAY         *psa;
		SAFEARRAYBOUND    rgsabound[1];

		rgsabound[0].lLbound = 0;
		rgsabound[0].cElements = srcDataBytes;
		psa = SafeArrayCreate(VT_UI1, 1, rgsabound);
		for (long i = 0; i < srcDataBytes; i++) {
			SafeArrayPutElement (psa, &i, srcData++);
		}

		varBLOB.vt = VT_ARRAY | VT_UI1;
		varBLOB.parray = psa;
		pDBRecordSet->GetFields()->GetItem(oleFiledName.GetBuffer())->AppendChunk(varBLOB);
		pDBRecordSet->Update();
		closeDB();
		return true;
	} catch(_com_error e) {
		AfxMessageBox("TZYDB::saveOleData 失败\n" + CString((LPCTSTR)e.Source()) + TEXT("\n描述:\n ") + CString((LPCTSTR)e.Description()));
	}
	closeDB();
	return false;
}
//------------------------------------------------------------------------
//调用示例：TZYDB::saveOleFromFile("select photo from BooksInfo where ID=1", "c:\\111.png", "photo")
bool TZYDB::saveOleFromFile(CString sSql, CString sFilePath, CString oleFiledName) {
	long fSize;

	FILE *fp = NULL;
	try{
		fopen_s(&fp, (LPCTSTR)sFilePath, "rb"); 
		
		if(!fp){
			throw CString("打开文件【" + sFilePath + "】失败");		  
		}

		fseek(fp, 0L, SEEK_END);
		fSize = ftell(fp);
		
		if(fSize == 0){
			throw CString("文件【" + sFilePath + "】字节数为0");
		}

		byte *pData = new byte[fSize];
		//byte *pData = new byte[fSize + 1];

		fseek(fp, 0L, SEEK_SET);

		fread(pData, 1, fSize, fp);

		fclose(fp);

		if(!pData){
			return false;
		}

		bool bOK = saveOleData(sSql, pData, fSize, oleFiledName);

		delete[] pData;

		return bOK;
	}catch (CString &e){
		AfxMessageBox(e); 
	} catch(...){
		AfxMessageBox("TZYDB::saveOleFromFile 失败!");
	}

	if(fp){
		fclose(fp);
	}

	return false;
}
//------------------------------------------------------------------------ 
long TZYDB::getOleBytes(CString sSql, CString oleFiledName){
	try {
		if(!OnInitADOConn()){
			return 0;
		}
		
		_variant_t _vl;
		_vl=(_variant_t)(LPCTSTR) sSql;
		
		pDBRecordSet->Open( _vl, _variant_t((IDispatch*)pDBConnection,true),
			adOpenStatic, adLockOptimistic, adCmdText);
		
		long dataBytes = pDBRecordSet->GetFields()->GetItem(oleFiledName.GetBuffer())->ActualSize;
		
		closeDB();
		
		return dataBytes;
	} catch (CString &eStr){
		AfxMessageBox("TZYDB::getOleBytes 失败\n" + eStr );
	} catch (_com_error e) {
		AfxMessageBox("TZYDB::getOleBytes失败\n" + CString((LPCTSTR)e.Source()) + TEXT("\n描述:\n ") + CString((LPCTSTR)e.Description()));
	} 	

	closeDB();
	return 0;
}
//------------------------------------------------------------------------ 
byte *TZYDB::oleToBufferByOpenDB(CString oleFiledName, long &dataBytes){ 
	dataBytes = pDBRecordSet->GetFields()->GetItem(oleFiledName.GetBuffer())->ActualSize;

	if(dataBytes < 1){
		return NULL;//throw CString("ole 数据长度为 0 ");
	}

	_variant_t    varBLOB;
	varBLOB = pDBRecordSet->GetFields()->GetItem(oleFiledName.GetBuffer())->GetChunk(dataBytes);

	if(varBLOB.vt != (VT_ARRAY | VT_UI1)){
		throw CString("ole 数据类型异常!!!");
	}

	byte *pData = new byte[dataBytes]; //主调函数释放
	
	if(!pData){
		throw CString("申请内存失败!!!");
	}

	byte *pBuf = NULL;
	SafeArrayAccessData(varBLOB.parray, (void **)&pBuf);
	memcpy(pData, pBuf, dataBytes);
	SafeArrayUnaccessData (varBLOB.parray);
	pBuf = 0;	

	return pData;
}
//------------------------------------------------------------------------
byte *TZYDB::readOleData(CString sSql, CString oleFiledName, long &dataBytes) { 
	try {
		if(!OnInitADOConn()){
			return NULL;
		}

		_variant_t _vl;
		_vl=(_variant_t)(LPCTSTR) sSql;
		
		pDBRecordSet->Open( _vl, _variant_t((IDispatch*)pDBConnection,true),
			adOpenStatic, adLockOptimistic, adCmdText);
		byte *pData = oleToBufferByOpenDB(oleFiledName, dataBytes);
		
		closeDB();
		
		return pData;
	} catch (CString &eStr){
		AfxMessageBox("TZYDB::ReadOleData 失败\n" + eStr );
	} catch (_com_error e) {
		AfxMessageBox("TZYDB::ReadOleData 失败\n" + CString((LPCTSTR)e.Source()) + TEXT("\n描述:\n ") + CString((LPCTSTR)e.Description()));
	} 		
	closeDB();
	return NULL;
}
//------------------------------------------------------------------------

 //第三周新增 end	  
//-------------------------------------------------------------
 //第八周 放在 Oracle.cpp中
//强制断开用户 sUserName 的连接(用户sUserName的客户端连接可能有若干个，全部强制断开连接)
bool TZYDB::disconnectUser(CString sUserName, bool m_bShowMes){
	struct LinkInfo{
		CString sid;
		CString serial;
	};

	CString strConnection;
	
	//system 密码不是 123456请自行修改
	strConnection = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(CID=GTU_APP)"
		"(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=127.0.0.1)(PORT=1521)))"
		"(CONNECT_DATA=(SID=orcl)(SERVER=DEDICATED)));User Id=system;Password=123456;PLSQLRSet=1;";	

	try {		
		if(FAILED(::CoInitialize(NULL))){
			throw CString("类工厂初始化失败！");		
		}

		_ConnectionPtr  pDBConnection; 
		_RecordsetPtr   pDBRecordSet;

		pDBConnection.CreateInstance("ADODB.Connection");

		_bstr_t bstrConnect = _T(strConnection.GetBuffer());
		
		//注意有的机子，若不是用管理员身份运行，此处可能无法建立连接
		pDBConnection->Open(bstrConnect, "","",adModeUnknown);		
		pDBRecordSet.CreateInstance("ADODB.Recordset");

		CString sSql;
		
		//用户名必须是大写
		sUserName = sUserName.MakeUpper();

		sSql = "select sid,serial# from v$session where username='" + sUserName + "'";

		pDBRecordSet->Open(
			sSql.GetBuffer(),
			pDBConnection.GetInterfacePtr(),
			adOpenDynamic, 
			adLockOptimistic, 
			adCmdText
			);

		CArray <LinkInfo> aLinkInfo;
		LinkInfo lnkInfo;

		while(!pDBRecordSet->adoEOF) {
			lnkInfo.sid = pDBRecordSet->GetCollect(_variant_t("sid"));
			lnkInfo.serial = pDBRecordSet->GetCollect(_variant_t("serial#"));

			aLinkInfo.Add(lnkInfo);
			
			pDBRecordSet->MoveNext();
		}

		pDBRecordSet->Close();
		pDBRecordSet.Release();
		pDBRecordSet = NULL;
		sSql = "";
		
		//遍历，断开用户的每一个连接
		for(int i = 0; i < aLinkInfo.GetSize(); i++){
			sSql = "alter system kill session '" 
				+ aLinkInfo[i].sid 
				+ "," + aLinkInfo[i].serial + "'";

			pDBConnection->Execute(sSql.GetBuffer(), NULL, 0);
		}
		
		if(m_bShowMes){
			sSql.Format("成功断开用户：%s 的  %d  个连接", sUserName.GetBuffer(), aLinkInfo.GetSize());
			AfxMessageBox(sSql);
		}

		if(pDBConnection->GetState() == adStateOpen){
			pDBConnection->Close();
		}

		pDBConnection.Release();
		pDBConnection = NULL;

		return true;
	}catch(_com_error e){
		AfxMessageBox("导入数据失败\n" + CString((LPCTSTR)e.Source()) +
			TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));	 		
	}

	return false;
}
//第八周End

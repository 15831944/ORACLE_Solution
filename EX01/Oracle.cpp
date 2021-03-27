#include "pch.h"
#include "Oracle.h"
#include "MyControls.h"
//------------------------------------------------------------------------
_ConnectionPtr    TZYDB::pDBConnection = NULL; // ���ݿ� 
_RecordsetPtr     TZYDB::pDBRecordSet  = NULL; // ���� 
_CommandPtr       TZYDB::pDBCommand = NULL; // ��¼
//------------------------------------------------------------------------
//����������ͼ���� start	  
bool TZYDB::loadImgFromFile(CDC *dstDC, CRect &dstRec,  CString sFile){
	try{
		CImage img;
		if(!img.IsNull()){
			img.Destroy();
		}

		img.Load(sFile);	   
		
		if (img.IsNull()){
			throw CString("����ͼ��" + sFile + "  ʧ��");
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
		s.Format("GlobalAlloc(GMEM_MOVEABLE, %d) ʧ��", dataBytes);
		throw CString(sError + s);	
	}

	void *pData = GlobalLock(hGlobal);
	memcpy(pData, pMemData, dataBytes);
	GlobalUnlock(hGlobal);

	if(CreateStreamOnHGlobal(hGlobal, TRUE,  &pStream) != S_OK){
		throw CString(sError + "CreateStreamOnHGlobalʧ��");
	}

	CImage *pImg = new CImage();
	if(pImg == NULL){
		sFlag = "pImg == NULL\r\nnew CImage()ʧ��";
	}else if (SUCCEEDED(pImg->Load(pStream))){
		if (pImg->IsNull()){
			sFlag = "��������ͼ��Ϊ��";
		}
	} else {
		sFlag = "pImg->Load(pStream)ʧ��";
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
		throw CString("CImage *ZYFunction::dbImgToCImage\r\n��ȡOle����ʧ��\n" + sSql);
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
		AfxMessageBox("TZYDB::showDBImg:  δ֪����");	         
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
	 //����������ͼ���� end 
//--------------------------------------------------------------------------
//��5���������� 
//--------------------------------------------------------------------------
int TZYDB::getRecordCount(CString sSql) {
	try {
		if(openDB(sSql) == NULL) return -1;
		_variant_t vIndex = (long)0;
		_variant_t vCount = pDBRecordSet->GetCollect(vIndex);
		closeDB();
		return vCount.lVal;
	}catch(_com_error e){		
		AfxMessageBox("TZYDB::getRecordCount ʧ��\n" + CString((LPCTSTR)e.Source()) + TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));//+ e.ErrorInfo());
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
		AfxMessageBox("TZYDB::getOneItem ʧ��\n" + CString((LPCTSTR)e.Source()) + TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));//+ e.ErrorInfo());
		closeDB();
	}
	return "";
}
//--------------------------------------------------------------------------
//��5���������� End
//--------------------------------------------------------------------------

//--------------------------------------------------------------------------
//������ Start
void TZYDB::dataSetToListCtrl(CListCtrl *pListCtrl, _RecordsetPtr pRecSet){
	try {		
		pListCtrl->SetExtendedStyle(pListCtrl->GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES );   
		
	
		CStringArray aFiledName;

		//��ȡ���ݼ����ֶ���
	    getFiledName(pRecSet, aFiledName);
		
		CString s = "";
		int iCount = aFiledName.GetSize();
	

		//�ռ���ʱ��Ҫ���ƣ��������������ʱ������
		pListCtrl->SetRedraw(FALSE);
		
		//��ձ��ؼ�
		clearCListCtrl(pListCtrl);

		//��д���ؼ���ͷ
		doListHeader(pListCtrl, aFiledName);

		//�������ݼ����Ʊ�
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
		AfxMessageBox("TZYDB::dataToListTable ʧ��\n" + eStr);
	}catch(_com_error e){		
		AfxMessageBox("TZYDB::dataToListTable ʧ��\n" + CString((LPCTSTR)e.Source()) + TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));//+ e.ErrorInfo());
		closeDB();
	}

	pListCtrl->SetRedraw(TRUE);
}
//--------------------------------------------------------------------------
/*���ڴ洢���̻�ȡ���ݼ�,sCommandTextӦ�����洢����������������������������һ������α�
	    �磺create or replace procedure ComHisStorage_2(sys_cur out sys_refcursor, byDate in date) 
		��sCommandTextӦΪ��ComHisStorage_2(to_date('2020-03-03','yyyy-mm-dd'))
*/
_RecordsetPtr TZYDB::getpRecSetByStoredProc(CString sCommandText){
	try{
		if(!TZYDB::OnInitADOConn()){
			return NULL;  
		}

		_CommandPtr m_pCommand;

		m_pCommand.CreateInstance(__uuidof(Command));		
		
		m_pCommand->ActiveConnection = TZYDB::pDBConnection;

		//�洢���̼�����
		m_pCommand->CommandText = sCommandText.AllocSysString(); 
		
		m_pCommand->CommandType = adCmdStoredProc;
		

		if (pDBRecordSet && pDBRecordSet->GetState() == adStateOpen){
			pDBRecordSet->Close();
		}
	
        TZYDB::pDBRecordSet = m_pCommand->Execute(NULL,NULL,adCmdStoredProc | adCmdUnspecified);	
     
		//���ص����ݼ�Ӧ�������������رգ����ͷ��ڴ�
		return TZYDB::pDBRecordSet;	
	}catch (_com_error &e){
			AfxMessageBox("" + CString((LPCTSTR)e.Source()) +
				TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));	  
	}

	return NULL;
}
//������End
//--------------------------------------------------------------------------

///////////////////////�ؼ��������������ݿ���ֱ�ӹ�ϵ��
void TZYDB::clearCListCtrl(CListCtrl *pListCtrl) {
	int nColumnCount = 0;
	//�Ӹ��ж�
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

	//InsertColumn �����麯������Ҫ�ֹ���̬����ָ������ת��
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
	
	//��ȡ�������
	CHeaderCtrl* pHeaderCtrl = pListCtrl->GetHeaderCtrl();
	if(!pHeaderCtrl){//�ޱ�ͷ
		return;
	}
	nColumnCount = pHeaderCtrl->GetItemCount();
	if(nColumnCount < 1){//���б�
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
///////////////////////�����������ݿ�ֱ����صĺ���
bool TZYDB::OnInitADOConn(CString linkStr) {//����ʼ��������
	static bool isFirst = true;//�ö���

	if(linkStr.IsEmpty()){
		linkStr = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(CID=GTU_APP)"
			"(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=127.0.0.1)(PORT=1521)))"
			"(CONNECT_DATA=(SID=orcl)(SERVER=DEDICATED)));"
			"User Id=student;Password=123456;PLSQLRSet=1;";
	}

	if(!isFirst){//ֻ��Ҫ����һ��
		return true;
	}

	isFirst = false;	
	
	try {		
		if(FAILED(::CoInitialize(NULL))){
			throw CString("�๤����ʼ��ʧ�ܣ�");
	    }
	
		pDBConnection.CreateInstance("ADODB.Connection");

		_bstr_t bstrConnect = _T(linkStr.GetBuffer());

		//2020-04-22 ���� 
		//ʹ�� �ͻ����α꣬����x64�ĳ��� �������ݼ�����ֻ�ܷ��ʵ���һ����¼
		//�����α귽ʽ ���� �� pDBConnection->Open ֮ǰִ��
		//����� �����ε� OnInitADOConn Ҳ��ͬ������
		pDBConnection->CursorLocation = adUseClient;
		//2020-04-22 ���� End

		//ע���еĻ��ӣ��������ù���Ա������У��˴������޷���������
		pDBConnection->Open(bstrConnect, "","",adModeUnknown);

		pDBRecordSet.CreateInstance("ADODB.Recordset");
		
		pDBCommand.CreateInstance(__uuidof(Command));
		
		pDBCommand->ActiveConnection = pDBConnection;
		
		return true;
	}catch(_com_error e){
		AfxMessageBox("AdoAccess::OnInitADOConn ����ʧ��\n" + CString((LPCTSTR)e.Source()) +
			TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));	 		
	}catch(CString &eStr){
		AfxMessageBox("AdoAccess::OnInitADOConn ����ʧ��\n" + eStr);
	}
	
	return false;
}

//------------------------------------------------------------------------
void TZYDB::closeConnetion(){//������������ã��ر�����
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
	    AfxMessageBox("TZYDB::closeConnetion\r\nδ֪����");
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
			AfxMessageBox("TZYDB::closeDB()\r\nδ֪����");
		}
	}	
}
//------------------------------------------------------------------------
_RecordsetPtr TZYDB::openDB(CString sSql, bool bShowErr) {
	try {
		if(!OnInitADOConn()){
			return NULL;
		}
		//2020-4-28����
		if(pDBRecordSet->GetState() == adStateOpen){
			pDBRecordSet->Close();
		}
		//2020-4-28���� End
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
			AfxMessageBox("TZYDB::openDB ʧ��\n" + CString((LPCTSTR)e.Source()) + TEXT("\nDescription:\n ") + 
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
		//2020-4-28����
		if(pDBRecordSet->GetState() == adStateOpen){
			pDBRecordSet->Close();
		}
		//2020-4-28���� End
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
			AfxMessageBox("TZYDB::openDB ʧ��\n" + CString((LPCTSTR)e.Source()) + TEXT("\nDescription:\n ") + 
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

		//��ȡ���ݼ����ֶ���
	    getFiledName(pDBRecordSet, aFiledName);
		
		CString s = "";
		int iCount = aFiledName.GetSize();
	

		//�ռ���ʱ��Ҫ���ƣ��������������ʱ������
		pListCtrl->SetRedraw(FALSE);
		
		//��ձ��ؼ�
		clearCListCtrl(pListCtrl);

		//��д���ؼ���ͷ
		doListHeader(pListCtrl, aFiledName);

		//�������ݼ����Ʊ�
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
		AfxMessageBox("TZYDB::dataToListTable ʧ��\n" + eStr);
	}catch(_com_error e){		
		AfxMessageBox("TZYDB::dataToListTable ʧ��\n" + CString((LPCTSTR)e.Source()) + TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));//+ e.ErrorInfo());
		closeDB();
	}

	pListCtrl->SetRedraw(TRUE);
	
	return false;
}
//-------------------------------------------------------------
 //���������� start	  
 // ����ɹ�������Ӱ��������������-1
//------------------------------------------------------------------------
int TZYDB::executeSql(CString sSql) {
	try {
		if(!OnInitADOConn()){
			return -1;  
		}

		_variant_t recordsAffected[800];
		
		//ִ��д���� ע:���Ƕ���pDBRecordSet->Open
		pDBConnection->Execute(sSql.GetBuffer(), recordsAffected, 0);
		char buffer[200];
		
		sprintf_s(buffer, "%s", (char*)((_bstr_t)recordsAffected));
		
		int iRecs = atoi(buffer);
		
		//closeConnetion();

		return iRecs;
	}catch(_com_error e){		
		AfxMessageBox("TZYDB::executeSql ʧ��\n" + CString((LPCTSTR)e.Source()) + TEXT("\n����:\n ") + CString((LPCTSTR)e.Description()));//+ e.ErrorInfo());
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
		s = "ִ������ʧ�ܣ������ѻ��ˡ�\r\n";
		s += CString((LPCTSTR)e.Source()) + "\r\n��������:\r\n" + CString((LPCTSTR)e.Description());

		//MessageBoxA(NULL,"�������ʧ��...","", MB_OK);
		MessageBoxA(NULL,s,"�������ʧ��...",MB_OK);
		//closeConnetion();

		return 0;
	}
}
//------------------------------------------------------------------------
//����ʾ����TZYDB::saveOleData("select photo from BooksInfor where ID=XXX", srcData, �ֽ���, "photo")
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
		AfxMessageBox("TZYDB::saveOleData ʧ��\n" + CString((LPCTSTR)e.Source()) + TEXT("\n����:\n ") + CString((LPCTSTR)e.Description()));
	}
	closeDB();
	return false;
}
//------------------------------------------------------------------------
//����ʾ����TZYDB::saveOleFromFile("select photo from BooksInfo where ID=1", "c:\\111.png", "photo")
bool TZYDB::saveOleFromFile(CString sSql, CString sFilePath, CString oleFiledName) {
	long fSize;

	FILE *fp = NULL;
	try{
		fopen_s(&fp, (LPCTSTR)sFilePath, "rb"); 
		
		if(!fp){
			throw CString("���ļ���" + sFilePath + "��ʧ��");		  
		}

		fseek(fp, 0L, SEEK_END);
		fSize = ftell(fp);
		
		if(fSize == 0){
			throw CString("�ļ���" + sFilePath + "���ֽ���Ϊ0");
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
		AfxMessageBox("TZYDB::saveOleFromFile ʧ��!");
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
		AfxMessageBox("TZYDB::getOleBytes ʧ��\n" + eStr );
	} catch (_com_error e) {
		AfxMessageBox("TZYDB::getOleBytesʧ��\n" + CString((LPCTSTR)e.Source()) + TEXT("\n����:\n ") + CString((LPCTSTR)e.Description()));
	} 	

	closeDB();
	return 0;
}
//------------------------------------------------------------------------ 
byte *TZYDB::oleToBufferByOpenDB(CString oleFiledName, long &dataBytes){ 
	dataBytes = pDBRecordSet->GetFields()->GetItem(oleFiledName.GetBuffer())->ActualSize;

	if(dataBytes < 1){
		return NULL;//throw CString("ole ���ݳ���Ϊ 0 ");
	}

	_variant_t    varBLOB;
	varBLOB = pDBRecordSet->GetFields()->GetItem(oleFiledName.GetBuffer())->GetChunk(dataBytes);

	if(varBLOB.vt != (VT_ARRAY | VT_UI1)){
		throw CString("ole ���������쳣!!!");
	}

	byte *pData = new byte[dataBytes]; //���������ͷ�
	
	if(!pData){
		throw CString("�����ڴ�ʧ��!!!");
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
		AfxMessageBox("TZYDB::ReadOleData ʧ��\n" + eStr );
	} catch (_com_error e) {
		AfxMessageBox("TZYDB::ReadOleData ʧ��\n" + CString((LPCTSTR)e.Source()) + TEXT("\n����:\n ") + CString((LPCTSTR)e.Description()));
	} 		
	closeDB();
	return NULL;
}
//------------------------------------------------------------------------

 //���������� end	  
//-------------------------------------------------------------
 //�ڰ��� ���� Oracle.cpp��
//ǿ�ƶϿ��û� sUserName ������(�û�sUserName�Ŀͻ������ӿ��������ɸ���ȫ��ǿ�ƶϿ�����)
bool TZYDB::disconnectUser(CString sUserName, bool m_bShowMes){
	struct LinkInfo{
		CString sid;
		CString serial;
	};

	CString strConnection;
	
	//system ���벻�� 123456�������޸�
	strConnection = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(CID=GTU_APP)"
		"(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=127.0.0.1)(PORT=1521)))"
		"(CONNECT_DATA=(SID=orcl)(SERVER=DEDICATED)));User Id=system;Password=123456;PLSQLRSet=1;";	

	try {		
		if(FAILED(::CoInitialize(NULL))){
			throw CString("�๤����ʼ��ʧ�ܣ�");		
		}

		_ConnectionPtr  pDBConnection; 
		_RecordsetPtr   pDBRecordSet;

		pDBConnection.CreateInstance("ADODB.Connection");

		_bstr_t bstrConnect = _T(strConnection.GetBuffer());
		
		//ע���еĻ��ӣ��������ù���Ա������У��˴������޷���������
		pDBConnection->Open(bstrConnect, "","",adModeUnknown);		
		pDBRecordSet.CreateInstance("ADODB.Recordset");

		CString sSql;
		
		//�û��������Ǵ�д
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
		
		//�������Ͽ��û���ÿһ������
		for(int i = 0; i < aLinkInfo.GetSize(); i++){
			sSql = "alter system kill session '" 
				+ aLinkInfo[i].sid 
				+ "," + aLinkInfo[i].serial + "'";

			pDBConnection->Execute(sSql.GetBuffer(), NULL, 0);
		}
		
		if(m_bShowMes){
			sSql.Format("�ɹ��Ͽ��û���%s ��  %d  ������", sUserName.GetBuffer(), aLinkInfo.GetSize());
			AfxMessageBox(sSql);
		}

		if(pDBConnection->GetState() == adStateOpen){
			pDBConnection->Close();
		}

		pDBConnection.Release();
		pDBConnection = NULL;

		return true;
	}catch(_com_error e){
		AfxMessageBox("��������ʧ��\n" + CString((LPCTSTR)e.Source()) +
			TEXT("\nDescription:\n ") + CString((LPCTSTR)e.Description()));	 		
	}

	return false;
}
//�ڰ���End
